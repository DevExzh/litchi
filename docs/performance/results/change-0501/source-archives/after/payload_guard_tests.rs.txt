use super::*;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, TargetMode};

use crate::{Error, Package, Result, SourceBackedPresentation, SourceBackedPresentationEditor};

type TestResult<T = ()> = Result<T>;

const CHART_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const CHART_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
const CHART_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";

#[derive(Clone, Copy, Debug)]
enum PayloadKind {
    Image,
    Chart,
}

#[test]
fn same_length_tampered_plan_payloads_reject_before_output() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = authored_fixture(&["destination-first", "destination-last"])?;

    for kind in [PayloadKind::Image, PayloadKind::Chart] {
        let source = open_source(&source_bytes)?;
        let editor = open_editor(&destination_bytes)?;
        let mut plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
        tamper_plan_payload(&mut plan, kind);
        let planned_digest = plan.touched_digest;

        let current = prepare(&editor, &source, 0, 1, 1, Some(&plan))?;
        assert_eq!(
            current.touched_digest, planned_digest,
            "metadata-only digest should remain unchanged for a same-length {kind:?} payload flip"
        );
        assert!(
            !current.matches(&plan),
            "exact prepared payload equality must reject a same-length {kind:?} payload flip"
        );

        let mut output = Vec::new();
        let error = editor
            .publish_cross_slide_copy_to_stream(&mut output, &plan)
            .expect_err("tampered private plan must not publish");
        assert!(matches!(error, Error::StaleSource));
        assert!(
            output.is_empty(),
            "stale plan must fail before writing output"
        );
    }

    Ok(())
}

#[test]
fn candidate_readback_rejects_same_length_image_and_chart_payloads() -> TestResult {
    let source_bytes = source_fixture()?;
    let destination_bytes = authored_fixture(&["destination-first", "destination-last"])?;

    for kind in [PayloadKind::Image, PayloadKind::Chart] {
        let source = open_source(&source_bytes)?;
        let editor = open_editor(&destination_bytes)?;
        let plan = editor.plan_cross_slide_copy(&source, 0, 1, 1)?;
        let mut candidate = prepare(&editor, &source, 0, 1, 1, Some(&plan))?;
        tamper_candidate_payload(&mut candidate, kind);

        let error = verify_candidate(&editor, &source, &candidate)
            .expect_err("candidate readback must reject changed payload bytes");
        assert!(matches!(error, Error::StaleSource));
    }

    Ok(())
}

fn tamper_plan_payload(plan: &mut SourceBackedCrossSlideCopyPlan, kind: PayloadKind) {
    match kind {
        PayloadKind::Image => {
            let image = plan
                .images
                .first_mut()
                .expect("source fixture should contain an image");
            image.bytes = flipped_payload(&image.bytes);
        },
        PayloadKind::Chart => {
            let chart = plan
                .charts
                .first_mut()
                .expect("source fixture should contain a chart");
            chart.bytes = flipped_payload(&chart.bytes);
        },
    }
}

fn tamper_candidate_payload(candidate: &mut Prepared, kind: PayloadKind) {
    match kind {
        PayloadKind::Image => {
            let image = candidate
                .images
                .first_mut()
                .expect("source fixture should contain an image");
            image.bytes = flipped_payload(&image.bytes);
        },
        PayloadKind::Chart => {
            let chart = candidate
                .charts
                .first_mut()
                .expect("source fixture should contain a chart");
            chart.bytes = flipped_payload(&chart.bytes);
        },
    }
}

fn flipped_payload(payload: &PreparedPayload) -> PreparedPayload {
    let mut bytes = payload.as_slice().to_vec();
    assert!(!bytes.is_empty(), "fixture payload must not be empty");
    bytes[0] ^= 0x5a;
    PreparedPayload::Owned(Arc::new(bytes))
}

fn source_fixture() -> TestResult<Vec<u8>> {
    let mut package = OpcPackage::from_vec(authored_fixture(&["source-one", "source-two"])?)?;
    let image_uri = PackURI::new("/ppt/media/guard-image.png").map_err(Error::Uri)?;
    let chart_uri = PackURI::new("/ppt/charts/guard-chart.xml").map_err(Error::Uri)?;
    package.try_add_part(Box::new(BlobPart::new(
        image_uri,
        ct::PNG.to_owned(),
        b"source-image-bytes".to_vec(),
    )))?;
    package.try_add_part(Box::new(BlobPart::new(
        chart_uri,
        CHART_CONTENT_TYPE.to_owned(),
        chart_payload(),
    )))?;

    let slide_uri = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Uri)?;
    let slide = package.get_part_mut(&slide_uri)?;
    slide.rels_mut().try_add_relationship(
        rt::IMAGE.to_owned(),
        "../media/guard-image.png".to_owned(),
        "rIdGuardImage".to_owned(),
        TargetMode::Internal,
    )?;
    slide.rels_mut().try_add_relationship(
        CHART_RELATIONSHIP_TYPE.to_owned(),
        "../charts/guard-chart.xml".to_owned(),
        "rIdGuardChart".to_owned(),
        TargetMode::Internal,
    )?;

    let xml = std::str::from_utf8(slide.blob())
        .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
    let insertion = xml
        .rfind("</p:spTree>")
        .ok_or_else(|| Error::Invalid("source fixture has no shape tree".into()))?;
    let graphics = r#"<p:pic><p:nvPicPr><p:cNvPr id="501" name="Guard Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdGuardImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="502" name="Guard Chart"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart r:id="rIdGuardChart"/></a:graphicData></a:graphic></p:graphicFrame>"#;
    let rewritten = format!("{}{graphics}{}", &xml[..insertion], &xml[insertion..]);
    slide.set_blob(rewritten.into_bytes());

    Ok(PackageWriter::to_bytes(&package)?)
}

fn chart_payload() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="{CHART_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:chart><c:autoTitleDeleted val="1"/></c:chart></c:chartSpace>"#
    )
    .into_bytes()
}

fn authored_fixture(slide_titles: &[&str]) -> TestResult<Vec<u8>> {
    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        for title in slide_titles {
            let slide = presentation.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(&format!("body:{title}"), 10, 20, 300, 400);
        }
    }
    let mut archive = OpcPackage::from_vec(package.to_bytes()?)?;
    for (index, title) in slide_titles.iter().enumerate() {
        let uri =
            PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).map_err(Error::Uri)?;
        let part = archive.get_part_mut(&uri)?;
        let xml = std::str::from_utf8(part.blob())
            .map_err(|error| Error::Invalid(format!("slide XML is not UTF-8: {error}")))?;
        let start = xml
            .find("name=\"")
            .ok_or_else(|| Error::Invalid("missing fixture slide name".into()))?
            + 6;
        let end = start
            + xml[start..]
                .find('"')
                .ok_or_else(|| Error::Invalid("unterminated fixture slide name".into()))?;
        let rewritten = format!("{}{title}{}", &xml[..start], &xml[end..]);
        part.set_blob(rewritten.into_bytes());
    }
    Ok(PackageWriter::to_bytes(&archive)?)
}

fn open_source(bytes: &[u8]) -> TestResult<SourceBackedPresentation> {
    SourceBackedPresentation::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
}

fn open_editor(bytes: &[u8]) -> TestResult<SourceBackedPresentationEditor> {
    SourceBackedPresentationEditor::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec())))
}
