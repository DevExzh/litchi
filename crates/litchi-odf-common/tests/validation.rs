#![allow(
    clippy::unwrap_used,
    reason = "fixed ZIP fixtures keep assertions focused on validation outcomes"
)]

use std::io::{Cursor, Write};

use litchi_core::{CheckStatus, IssueSeverity, ValidationLimitKind, ValidationLimits};
use litchi_odf_common::{
    OdfValidationError, OdfValidationLimits, validate_package, validate_package_with_limits,
};
use zip::{
    CompressionMethod, ZipWriter,
    write::{FullFileOptions, SimpleFileOptions},
};

const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const OFFICE_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";

fn manifest(mimetype: &str, extra_entries: &str) -> String {
    format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>{extra_entries}</manifest:manifest>"
    )
}

fn content(family: &str, extra_namespace: &str, body: &str) -> String {
    let extra_namespace = if extra_namespace.is_empty() {
        String::new()
    } else {
        format!(" {extra_namespace}")
    };
    format!(
        "<office:document-content xmlns:office=\"{OFFICE_NS}\"{extra_namespace}><office:body><office:{family}>{body}</office:{family}></office:body></office:document-content>"
    )
}

fn build_package(
    mimetype: &str,
    manifest_xml: &[u8],
    content_xml: &[u8],
    extras: &[(&str, &[u8])],
) -> Vec<u8> {
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    zip.write_all(manifest_xml).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    zip.write_all(content_xml).unwrap();
    for (path, bytes) in extras {
        zip.start_file(*path, deflated).unwrap();
        zip.write_all(bytes).unwrap();
    }
    zip.finish().unwrap().into_inner()
}

fn valid_package(mimetype: &str, family: &str) -> Vec<u8> {
    build_package(
        mimetype,
        manifest(mimetype, "").as_bytes(),
        content(family, "", "").as_bytes(),
        &[],
    )
}

fn status<'a>(report: &'a litchi_core::ValidateReport, id: &str) -> &'a CheckStatus {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .unwrap()
        .status()
}

fn has_code(report: &litchi_core::ValidateReport, code: &str) -> bool {
    report.issues().iter().any(|issue| issue.code() == code)
}

#[test]
fn accepts_minimal_text_spreadsheet_and_presentation_packages() {
    let fixtures = [
        ("application/vnd.oasis.opendocument.text", "text"),
        (
            "application/vnd.oasis.opendocument.spreadsheet",
            "spreadsheet",
        ),
        (
            "application/vnd.oasis.opendocument.presentation",
            "presentation",
        ),
        ("application/vnd.oasis.opendocument.text-template", "text"),
    ];
    for (mimetype, family) in fixtures {
        let report = validate_package(&valid_package(mimetype, family)).unwrap();
        assert!(report.is_complete());
        assert!(!report.has_errors());
        assert!(report.issues().is_empty());
        assert!(matches!(
            status(&report, "odf.content_xml.external_reference_presence"),
            CheckStatus::NotApplicable
        ));
    }
}

#[test]
fn malformed_zip_stops_every_dependent_capability_without_panicking() {
    let report = validate_package(b"not a zip").unwrap();
    assert!(report.has_fatal());
    assert!(has_code(&report, "odf.zip.invalid"));
    assert!(matches!(
        status(&report, "odf.package.ingress"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "odf.package.catalog"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn malformed_manifest_and_xml_are_content_free_deterministic_issues() {
    let secret = "customer-secret-638ae7";
    let malformed_manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"{secret}\"/><manifest:file-entry manifest:full-path=\"{secret}\"/></manifest:manifest>"
    );
    let bytes = build_package(
        "application/vnd.oasis.opendocument.text",
        malformed_manifest.as_bytes(),
        b"<office:document-content>",
        &[],
    );
    let first = validate_package(&bytes).unwrap();
    let second = validate_package(&bytes).unwrap();
    assert_eq!(
        first
            .issues()
            .iter()
            .map(|issue| issue.id())
            .collect::<Vec<_>>(),
        second
            .issues()
            .iter()
            .map(|issue| issue.id())
            .collect::<Vec<_>>()
    );
    assert!(has_code(&first, "odf.manifest.invalid"));
    for issue in first.issues() {
        assert!(!issue.message().contains(secret));
        assert!(!issue.code().contains(secret));
        assert!(issue.locations().iter().all(|location| {
            !location.part().unwrap_or_default().contains(secret)
                && !location.path().unwrap_or_default().contains(secret)
        }));
    }

    let valid_manifest = manifest("application/vnd.oasis.opendocument.text", "");
    let malformed_xml = build_package(
        "application/vnd.oasis.opendocument.text",
        valid_manifest.as_bytes(),
        b"<office:document-content>",
        &[],
    );
    let report = validate_package(&malformed_xml).unwrap();
    assert!(has_code(&report, "odf.xml.malformed"));
    assert!(matches!(
        status(&report, "odf.package.root_xml"),
        CheckStatus::Complete
    ));
    assert!(matches!(
        status(&report, "odf.content_xml.external_reference_presence"),
        CheckStatus::Blocked { .. }
    ));

    let fake_root = format!(
        "<fake xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></fake>"
    );
    let fake = build_package(
        "application/vnd.oasis.opendocument.text",
        fake_root.as_bytes(),
        content("text", "", "").as_bytes(),
        &[],
    );
    assert!(has_code(
        &validate_package(&fake).unwrap(),
        "odf.manifest.root_invalid"
    ));

    let misplaced = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:encryption-data/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"
    );
    let misplaced_bytes = build_package(
        "application/vnd.oasis.opendocument.text",
        misplaced.as_bytes(),
        content("text", "", "").as_bytes(),
        &[],
    );
    assert!(has_code(
        &validate_package(&misplaced_bytes).unwrap(),
        "odf.manifest.encryption_placement_invalid"
    ));

    let duplicate = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"><manifest:encryption-data/><manifest:encryption-data/></manifest:file-entry></manifest:manifest>"
    );
    let duplicate_bytes = build_package(
        "application/vnd.oasis.opendocument.text",
        duplicate.as_bytes(),
        b"ciphertext",
        &[],
    );
    assert!(has_code(
        &validate_package(&duplicate_bytes).unwrap(),
        "odf.manifest.encryption_placement_invalid"
    ));
}

#[test]
fn reports_mimetype_catalog_and_hostile_path_failures() {
    let declared = "application/vnd.oasis.opendocument.text";
    let extra = "<manifest:file-entry manifest:full-path=\"missing.bin\" manifest:media-type=\"application/octet-stream\"/><manifest:file-entry manifest:full-path=\"evil.bin\" manifest:media-type=\"application/octet-stream\"/>";
    let bytes = build_package(
        "application/vnd.oasis.opendocument.spreadsheet",
        manifest(declared, extra).as_bytes(),
        content("text", "", "").as_bytes(),
        &[("../evil.bin", b"hostile")],
    );
    let report = validate_package(&bytes).unwrap();
    assert!(has_code(&report, "odf.mimetype.inconsistent"));
    assert!(has_code(&report, "odf.catalog.inconsistent"));
    assert!(has_code(&report, "odf.catalog.hostile_path"));
    assert!(report.has_errors());

    let unknown = "application/vnd.oasis.opendocument.attacker";
    let unknown_bytes = build_package(
        unknown,
        manifest(unknown, "").as_bytes(),
        content("text", "", "").as_bytes(),
        &[],
    );
    assert!(has_code(
        &validate_package(&unknown_bytes).unwrap(),
        "odf.mimetype.inconsistent"
    ));
}

#[test]
fn duplicate_normalized_member_names_are_an_ingress_rejection() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let extra = "<manifest:file-entry manifest:full-path=\"evil.bin\" manifest:media-type=\"application/octet-stream\"/>";
    let bytes = build_package(
        mimetype,
        manifest(mimetype, extra).as_bytes(),
        content("text", "", "").as_bytes(),
        &[("evil.bin", b"first"), ("../evil.bin", b"second")],
    );
    let report = validate_package(&bytes).unwrap();
    assert!(has_code(&report, "odf.zip.invalid"));
    assert!(report.has_fatal());
    assert!(matches!(
        status(&report, "odf.package.catalog"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn mimetype_local_header_extra_fields_are_a_layout_error() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let mut mimetype_options =
        FullFileOptions::default().compression_method(CompressionMethod::Stored);
    mimetype_options
        .add_extra_data(0x1234, b"extra", false)
        .unwrap();
    zip.start_file("mimetype", mimetype_options).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file(
        "META-INF/manifest.xml",
        SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(manifest(mimetype, "").as_bytes()).unwrap();
    zip.start_file(
        "content.xml",
        SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(content("text", "", "").as_bytes()).unwrap();
    let bytes = zip.finish().unwrap().into_inner();
    let report = validate_package(&bytes).unwrap();
    assert!(has_code(&report, "odf.mimetype.layout"));
}

#[test]
fn central_file_comments_are_charged_to_the_raw_metadata_ceiling() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let mut zip = ZipWriter::new(Cursor::new(Vec::new()));
    let commented = FullFileOptions::default()
        .compression_method(CompressionMethod::Stored)
        .with_file_comment("x".repeat(256));
    zip.start_file("mimetype", commented).unwrap();
    zip.write_all(mimetype.as_bytes()).unwrap();
    zip.start_file(
        "META-INF/manifest.xml",
        SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(manifest(mimetype, "").as_bytes()).unwrap();
    zip.start_file(
        "content.xml",
        SimpleFileOptions::default().compression_method(CompressionMethod::Deflated),
    )
    .unwrap();
    zip.write_all(content("text", "", "").as_bytes()).unwrap();
    let bytes = zip.finish().unwrap().into_inner();
    let report = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_archive_metadata_bytes(128),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "odf.package.ingress"),
        CheckStatus::Blocked { .. }
    ));
}

#[test]
fn reports_encryption_signature_external_and_macro_presence_only() {
    use litchi_odf_common::core::{PackageWriter, Profile};

    let mut encrypted_writer = PackageWriter::new();
    encrypted_writer
        .set_mimetype("application/vnd.oasis.opendocument.text")
        .unwrap();
    encrypted_writer
        .set_encryption("password", Profile::compatible())
        .unwrap();
    encrypted_writer
        .add_file("content.xml", content("text", "", "").as_bytes())
        .unwrap();
    let encrypted = encrypted_writer.finish().unwrap();
    let encrypted_report = validate_package(&encrypted).unwrap();
    assert!(has_code(
        &encrypted_report,
        "odf.encryption.infrastructure_present"
    ));
    assert!(matches!(
        status(&encrypted_report, "odf.package.root_xml"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(
            &encrypted_report,
            "odf.content_xml.external_reference_presence"
        ),
        CheckStatus::StoppedBy { .. }
    ));

    let unsupported_encryption_manifest = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"><manifest:encryption-data><manifest:algorithm manifest:algorithm-name=\"urn:example:unsupported\"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>"
    );
    let unsupported_encryption = build_package(
        "application/vnd.oasis.opendocument.text",
        unsupported_encryption_manifest.as_bytes(),
        b"ciphertext",
        &[],
    );
    let unsupported_report = validate_package(&unsupported_encryption).unwrap();
    assert!(has_code(
        &unsupported_report,
        "odf.encryption.infrastructure_present"
    ));
    assert!(!has_code(&unsupported_report, "odf.manifest.invalid"));
    assert!(matches!(
        status(&unsupported_report, "odf.package.root_xml"),
        CheckStatus::Blocked { .. }
    ));

    let mimetype = "application/vnd.oasis.opendocument.text";
    let extras = "<manifest:file-entry manifest:full-path=\"META-INF/documentsignatures.xml\" manifest:media-type=\"application/vnd.oasis.opendocument.digital-signature+xml\"/><manifest:file-entry manifest:full-path=\"Basic/module.xml\" manifest:media-type=\"text/xml\"/>";
    let linked = content(
        "text",
        "xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        "<office:a xlink:href=\"h&#116;tps://example.invalid/document\"/>",
    );
    let bytes = build_package(
        mimetype,
        manifest(mimetype, extras).as_bytes(),
        linked.as_bytes(),
        &[
            ("META-INF/documentsignatures.xml", b"<signatures/>"),
            ("Basic/module.xml", b"<macro/>"),
        ],
    );
    let report = validate_package(&bytes).unwrap();
    assert!(has_code(&report, "odf.signature.infrastructure_present"));
    assert!(has_code(&report, "odf.external_reference.present"));
    assert!(has_code(&report, "odf.macro.storage_present"));
    assert!(
        report
            .issues()
            .iter()
            .filter(|issue| issue.severity() >= IssueSeverity::Warning)
            .all(|issue| issue.repair().repair_id().is_none())
    );

    let whitespace_padded_link = content(
        "text",
        "xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        "<office:a xlink:href=\"  https://example.invalid/document  \"/>",
    );
    let whitespace_padded = build_package(
        mimetype,
        manifest(mimetype, "").as_bytes(),
        whitespace_padded_link.as_bytes(),
        &[],
    );
    assert!(has_code(
        &validate_package(&whitespace_padded).unwrap(),
        "odf.external_reference.present"
    ));

    let internal_links = content(
        "text",
        "xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
        "<office:a xlink:href=\" ./Pictures/a.png \"/><office:a xlink:href=\" #bookmark \"/>",
    );
    let internal = build_package(
        mimetype,
        manifest(mimetype, "").as_bytes(),
        internal_links.as_bytes(),
        &[],
    );
    let internal_report = validate_package(&internal).unwrap();
    assert!(matches!(
        status(
            &internal_report,
            "odf.content_xml.external_reference_presence"
        ),
        CheckStatus::NotApplicable
    ));
    assert!(!has_code(
        &internal_report,
        "odf.external_reference.present"
    ));

    for unsafe_href in [
        "/outside",
        "../../outside",
        r"\\server\outside",
        "Pictures/a.png?cache=1",
        "Pictures/a.png#embedded",
        "Pictures/a%3Fcache.png",
        "Pictures/a%23embedded.png",
    ] {
        let unsafe_link = content(
            "text",
            "xmlns:xlink=\"http://www.w3.org/1999/xlink\"",
            &format!("<office:a xlink:href=\"{unsafe_href}\"/>"),
        );
        let unsafe_package = build_package(
            mimetype,
            manifest(mimetype, "").as_bytes(),
            unsafe_link.as_bytes(),
            &[],
        );
        assert!(has_code(
            &validate_package(&unsafe_package).unwrap(),
            "odf.external_reference.present"
        ));
    }
}

#[test]
fn input_archive_manifest_xml_and_report_limits_are_distinguished() {
    let bytes = valid_package("application/vnd.oasis.opendocument.text", "text");

    let source = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_input_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&source, "odf.package.ingress"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&source, "odf.package.catalog"),
        CheckStatus::StoppedBy { .. }
    ));

    let entries = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_entries(2),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&entries, "odf.package.ingress"),
        CheckStatus::Blocked { .. }
    ));

    let entry_bytes = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_archive_entry_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&entry_bytes, "odf.package.ingress"),
        CheckStatus::Blocked { .. }
    ));

    let manifest_limit_report = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_manifest_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&manifest_limit_report, "odf.package.mimetype_manifest"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&manifest_limit_report, "odf.package.catalog"),
        CheckStatus::StoppedBy { .. }
    ));

    let manifest_entries = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_manifest_entries(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&manifest_entries, "odf.package.mimetype_manifest"),
        CheckStatus::Blocked { .. }
    ));

    let xml_bytes = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_root_xml_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&xml_bytes, "odf.package.root_xml"),
        CheckStatus::Blocked { .. }
    ));

    let xml_events = validate_package_with_limits(
        &bytes,
        OdfValidationLimits::default().with_max_xml_events(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&xml_events, "odf.package.root_xml"),
        CheckStatus::Blocked { .. }
    ));

    let report_limits = ValidationLimits::new(7, 32, 4, 4, 128, 512, 256, 256, 256, 16_384);
    let error = validate_package_with_limits(&bytes, OdfValidationLimits::default(), report_limits)
        .unwrap_err();
    assert!(matches!(
        error,
        OdfValidationError::Report(litchi_core::ValidationReportError::Limit {
            kind: ValidationLimitKind::Checks,
            ..
        })
    ));

    let hostile = build_package(
        "application/vnd.oasis.opendocument.text",
        manifest("application/vnd.oasis.opendocument.text", "").as_bytes(),
        content("text", "", "").as_bytes(),
        &[("../hostile.bin", b"x")],
    );
    let issue_limits = ValidationLimits::new(8, 0, 4, 4, 128, 512, 256, 256, 256, 16_384);
    let issue_error =
        validate_package_with_limits(&hostile, OdfValidationLimits::default(), issue_limits)
            .unwrap_err();
    assert!(matches!(
        issue_error,
        OdfValidationError::Report(litchi_core::ValidationReportError::Limit {
            kind: ValidationLimitKind::Issues,
            observed: 1,
            limit: 0,
        })
    ));
}

#[test]
fn doctype_and_depth_limit_do_not_fetch_or_expand_entities() {
    let mimetype = "application/vnd.oasis.opendocument.text";
    let doctype = b"<!DOCTYPE office [<!ENTITY leaked SYSTEM 'file:///secret'>]><office:document-content xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'><office:body>&leaked;</office:body></office:document-content>";
    let bytes = build_package(mimetype, manifest(mimetype, "").as_bytes(), doctype, &[]);
    let report = validate_package(&bytes).unwrap();
    assert!(has_code(&report, "odf.xml.doctype_forbidden"));

    let undeclared = b"<office:document-content xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'><office:body>&leaked;</office:body></office:document-content>";
    let undeclared_bytes =
        build_package(mimetype, manifest(mimetype, "").as_bytes(), undeclared, &[]);
    assert!(has_code(
        &validate_package(&undeclared_bytes).unwrap(),
        "odf.xml.malformed"
    ));

    let invalid_content_reference = b"<office:document-content xmlns:office='urn:oasis:names:tc:opendocument:xmlns:office:1.0'><office:body>&#0;</office:body></office:document-content>";
    let invalid_content_bytes = build_package(
        mimetype,
        manifest(mimetype, "").as_bytes(),
        invalid_content_reference,
        &[],
    );
    assert!(has_code(
        &validate_package(&invalid_content_bytes).unwrap(),
        "odf.xml.malformed"
    ));

    let invalid_manifest_reference = format!(
        "<manifest:manifest xmlns:manifest=\"{MANIFEST_NS}\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"{mimetype}\"/><manifest:file-entry manifest:full-path=\"content&#xD800;.xml\" manifest:media-type=\"text/xml\"/></manifest:manifest>"
    );
    let invalid_manifest_bytes = build_package(
        mimetype,
        invalid_manifest_reference.as_bytes(),
        content("text", "", "").as_bytes(),
        &[],
    );
    assert!(has_code(
        &validate_package(&invalid_manifest_bytes).unwrap(),
        "odf.manifest.invalid"
    ));

    let nested = content("text", "", "<office:a><office:b/></office:a>");
    let nested_bytes = build_package(
        mimetype,
        manifest(mimetype, "").as_bytes(),
        nested.as_bytes(),
        &[],
    );
    let limited = validate_package_with_limits(
        &nested_bytes,
        OdfValidationLimits::default().with_max_xml_depth(2),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&limited, "odf.package.root_xml"),
        CheckStatus::Blocked { .. }
    ));
}
