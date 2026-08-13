#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "validation fixtures intentionally panic on construction failure"
)]

use std::sync::Arc;

use litchi_core::{CheckStatus, OwnedSource, ReadAt, ValidationLimits};
use litchi_opc::SourceBackedPackage;
use litchi_pptx::{
    PptxValidationLimits, validate_source_backed, validate_source_backed_with_limits,
    validate_source_with_limits,
};
use soapberry_zip::office::StreamingArchiveWriter;

const PML: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SLIDE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const SIGNATURE_ORIGIN: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const VBA_SIGNATURE_AGILE: &str =
    "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile";

fn status<'a>(report: &'a litchi_core::ValidateReport, id: &str) -> &'a CheckStatus {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .expect("declared PPTX validation capability")
        .status()
}

fn package_bytes(
    presentation: &str,
    slide: &str,
    presentation_relationships: &str,
    package_relationships: &str,
    extra_members: &[(&str, &[u8])],
) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>{}</Types>"#,
        extra_members
            .iter()
            .filter(|(name, _)| name.ends_with("vbaProject.bin"))
            .map(|_| {
                "<Override PartName=\"/ppt/vbaProject.bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>"
            })
            .collect::<String>()
            + &extra_members
                .iter()
                .filter(|(name, _)| name.ends_with("origin.sigs"))
                .map(|_| "<Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/>")
                .collect::<String>()
            + &extra_members
                .iter()
                .filter(|(name, _)| name.ends_with("vbaProjectSignatureAgile.bin"))
                .map(|_| "<Override PartName=\"/ppt/vbaProjectSignatureAgile.bin\" ContentType=\"application/vnd.ms-office.vbaProjectSignatureAgile\"/>")
                .collect::<String>()
            + &extra_members
                .iter()
                .filter(|(name, _)| name.ends_with("orphan.xml"))
                .map(|_| "<Override PartName=\"/ppt/orphan.xml\" ContentType=\"application/xml\"/>")
                .collect::<String>()
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", package_relationships.as_bytes())
        .unwrap();
    writer
        .write_stored("ppt/presentation.xml", presentation.as_bytes())
        .unwrap();
    writer
        .write_stored(
            "ppt/_rels/presentation.xml.rels",
            presentation_relationships.as_bytes(),
        )
        .unwrap();
    writer
        .write_stored("ppt/slides/slide1.xml", slide.as_bytes())
        .unwrap();
    for (name, bytes) in extra_members {
        writer.write_stored(name, bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn valid_package() -> Vec<u8> {
    package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/><Relationship Id="rIdExternal" Type="{IMAGE}" Target="https://example.invalid/image.png" TargetMode="External"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    )
}

#[test]
fn valid_source_report_reuses_catalog_and_reports_external_presence_only() {
    let bytes = valid_package();
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at(source).unwrap();
    let report = validate_source_backed(&package).unwrap();

    assert!(matches!(
        status(&report, "pptx.package.loaded_relationships_content_types"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.package.relationship_graph"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.presentation.root"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.presentation.ordered_slide_closure"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.package.external_target_presence"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.package.signature_presence"),
        CheckStatus::NotApplicable
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.external_target.present")
    );
    assert!(!report.has_fatal());
}

#[test]
fn mce_is_reported_without_selecting_a_branch() {
    let presentation = format!(
        r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Fallback><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></mc:Fallback></mc:AlternateContent><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
    );
    let bytes = package_bytes(
        &presentation,
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();

    assert!(matches!(
        status(&report, "pptx.presentation.root"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.presentation.ordered_slide_closure"),
        CheckStatus::NotApplicable
    ));
    assert!(matches!(
        status(&report, "pptx.presentation.mce_presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.mce.present")
    );
}

#[test]
fn malformed_slide_is_a_bounded_issue_and_xml_limit_stops_dependents() {
    let bytes = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}">"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.presentation.ordered_slide_closure"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.slide_closure.incomplete")
    );

    let limited = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        PptxValidationLimits::default().with_max_xml_bytes(16),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&limited, "pptx.presentation.root"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&limited, "pptx.presentation.ordered_slide_closure"),
        CheckStatus::StoppedBy { check } if check.as_str() == "pptx.presentation.root"
    ));
}

#[test]
fn signature_and_macro_presence_are_inert_facts() {
    let bytes = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/><Relationship Id="rIdSig" Type="{SIGNATURE_ORIGIN}" Target="_xmlsignatures/origin.sigs"/></Relationships>"#
        ),
        &[
            ("_xmlsignatures/origin.sigs", b"opaque-signature".as_slice()),
            ("ppt/vbaProject.bin", b"opaque-vba".as_slice()),
        ],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.package.signature_presence"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "pptx.package.macro_presence"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.signature.infrastructure_present")
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.macro.storage_present")
    );
}

#[test]
fn source_backed_limits_are_finite_and_report_owned_xml_ceiling() {
    let bytes = valid_package();
    let package = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    let report = validate_source_backed_with_limits(
        &package,
        PptxValidationLimits::default().with_max_owner_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.presentation.root"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn strict_namespaces_and_direct_numeric_slide_owners_are_required() {
    let undeclared_relationship_prefix = package_bytes(
        r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#,
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(undeclared_relationship_prefix)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.root.malformed")
    );

    let duplicate_owners = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(duplicate_owners)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.presentation.root"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.slide_closure.incomplete")
    );

    let non_numeric = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="not-a-number" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(non_numeric)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.root.malformed")
    );
}

#[test]
fn malformed_references_comments_and_depth_block_mce_truthfully() {
    let unknown_reference = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}">bad &unknown;</p:sld>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(unknown_reference)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.slide_closure.incomplete")
    );
    assert!(matches!(
        status(&report, "pptx.presentation.mce_presence"),
        CheckStatus::Blocked { .. }
    ));

    let invalid_comment = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><!--invalid--comment--><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(invalid_comment)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.root.malformed")
    );

    let limited = validate_source_with_limits(
        Arc::new(OwnedSource::new(valid_package())),
        PptxValidationLimits::default().with_max_xml_depth(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&limited, "pptx.presentation.mce_presence"),
        CheckStatus::StoppedBy { check } if check.as_str() == "pptx.presentation.root"
    ));
}

#[test]
fn graph_checks_orphan_manifests_and_agile_vba_signature_presence() {
    let orphan_relationships = format!(
        r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdMissing" Type="{IMAGE}" Target="missing.bin"/></Relationships>"#
    );
    let package_relationships = format!(
        r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/><Relationship Id="rIdAgile" Type="{VBA_SIGNATURE_AGILE}" Target="ppt/vbaProjectSignatureAgile.bin"/></Relationships>"#
    );
    let bytes = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &package_relationships,
        &[
            ("ppt/orphan.xml", b"orphan".as_slice()),
            ("ppt/_rels/orphan.xml.rels", orphan_relationships.as_bytes()),
            (
                "ppt/vbaProjectSignatureAgile.bin",
                b"opaque-agile-signature".as_slice(),
            ),
        ],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.package.relationship_graph"),
        CheckStatus::Complete
    ));
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.relationship_graph.incomplete")
    );
    assert!(matches!(
        status(&report, "pptx.package.macro_presence"),
        CheckStatus::Complete
    ));
}

#[test]
fn mce_conclusion_is_order_independent_when_a_slide_cannot_be_inspected() {
    for presentation in [
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="256" r:id="rIdSlide"/><p:sldId id="257" r:id="rIdMissing"/></p:sldIdLst></p:presentation>"#
        ),
        format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="257" r:id="rIdMissing"/><p:sldId id="256" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
        ),
    ] {
        let bytes = package_bytes(
            &presentation,
            &format!(
                r#"<p:sld xmlns:p="{PML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent/></p:sld>"#
            ),
            &format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/><Relationship Id="rIdMissing" Type="{SLIDE}" Target="slides/missing.xml"/></Relationships>"#
            ),
            &format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
            ),
            &[],
        );
        let report = validate_source_with_limits(
            Arc::new(OwnedSource::new(bytes)),
            PptxValidationLimits::default(),
            ValidationLimits::default(),
        )
        .unwrap();
        assert!(matches!(
            status(&report, "pptx.presentation.mce_presence"),
            CheckStatus::Blocked { .. }
        ));
    }
}

#[test]
fn slide_id_schema_bounds_and_direct_list_cardinality_are_enforced() {
    for id in ["255", "2147483648"] {
        let bytes = package_bytes(
            &format!(
                r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst><p:sldId id="{id}" r:id="rIdSlide"/></p:sldIdLst></p:presentation>"#
            ),
            &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
            &format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
            ),
            &format!(
                r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
            ),
            &[],
        );
        let report = validate_source_with_limits(
            Arc::new(OwnedSource::new(bytes)),
            PptxValidationLimits::default(),
            ValidationLimits::default(),
        )
        .unwrap();
        assert!(
            report
                .issues()
                .iter()
                .any(|issue| issue.code() == "pptx.presentation.root.malformed")
        );
    }

    let multiple_lists = package_bytes(
        &format!(
            r#"<p:presentation xmlns:p="{PML}" xmlns:r="{REL}"><p:sldIdLst/><p:sldIdLst/></p:presentation>"#
        ),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(multiple_lists)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.root.malformed")
    );

    let no_list = package_bytes(
        &format!(r#"<p:presentation xmlns:p="{PML}"/>"#),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(no_list)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "pptx.presentation.root"),
        CheckStatus::Complete
    ));
}

#[test]
fn namespace_attribute_values_and_reader_depth_are_bounded() {
    let invalid_namespace_value = package_bytes(
        &format!("<p:presentation xmlns:p=\"{PML}\0\"><p:sldIdLst/></p:presentation>"),
        &format!(r#"<p:sld xmlns:p="{PML}"/>"#),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rIdSlide" Type="{SLIDE}" Target="slides/slide1.xml"/></Relationships>"#
        ),
        &format!(
            r#"<Relationships xmlns="{PKG_REL}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT}" Target="ppt/presentation.xml"/></Relationships>"#
        ),
        &[],
    );
    let report = validate_source_with_limits(
        Arc::new(OwnedSource::new(invalid_namespace_value)),
        PptxValidationLimits::default(),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "pptx.presentation.root.malformed")
    );

    let limits = PptxValidationLimits::default().with_max_xml_depth(usize::MAX);
    assert_eq!(limits.max_xml_depth(), u16::MAX as usize - 1);
}
