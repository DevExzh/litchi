//! Malformed-input hardening for format-neutral ODF parsers.

use litchi_core::Error;
use litchi_odf_common::core::metadata::Metadata as OdfMetadata;
use litchi_odf_common::package::{Archive, ArchiveLimits, parse_manifest, read_manifest};
use litchi_odf_common::{Error as OdfError, XmlStreamLimits, chart, drawing, embedded, media};
use std::io::{Cursor, Write as _};

const CHART_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:title><text:p>Revenue</text:p></chart:title><chart:legend chart:legend-position="end"/><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"/><chart:series chart:values-cell-range-address="Sheet1.$B$1:.$B$3"><chart:data-point chart:repeated="2"/></chart:series><table:table table:name="local-table"/></chart:plot-area></chart:chart></office:chart></office:body></office:document-content>"#;
const FLAT_MEDIA_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:math="http://www.w3.org/1998/Math/MathML" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.text" office:version="1.3"><office:body><office:text><draw:frame draw:name="image" svg:width="1cm" svg:height="1cm"><draw:image><office:binary-data>aQ==</office:binary-data></draw:image></draw:frame><draw:frame draw:name="formula" svg:width="1cm" svg:height="1cm"><draw:object><math:math><math:mi>x</math:mi></math:math></draw:object></draw:frame><text:p>body</text:p></office:text></office:body></office:document>"#;

fn assert_invalid<T>(result: Result<T, Error>, case: &str) {
    match result {
        Err(Error::InvalidFormat(_)) => {},
        Err(error) => panic!("{case} produced a non-format error: {error:?}"),
        Ok(_) => panic!("{case} unexpectedly parsed"),
    }
}

fn package() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut output = Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored)?;
    zip.write_all(b"application/vnd.oasis.opendocument.text")?;
    zip.start_file("content.xml", deflated)?;
    zip.write_all(b"<office:document-content/>")?;
    zip.start_file("META-INF/manifest.xml", deflated)?;
    zip.write_all(br#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#)?;
    zip.finish()?;
    Ok(output.into_inner())
}

fn exercise_archive(bytes: &[u8]) {
    if let Ok(archive) = Archive::new(bytes, ArchiveLimits::default()) {
        drop(read_manifest(&archive, XmlStreamLimits::default()));
    }
}

#[test]
fn chart_truncation_and_mutation_sweeps_never_panic() -> Result<(), Error> {
    for end in 0..CHART_SEED.len() {
        drop(chart_read(&CHART_SEED[..end]));
    }
    for position in 0..CHART_SEED.len() {
        let mut mutated = CHART_SEED.as_bytes().to_vec();
        mutated[position] ^= 1;
        if let Ok(xml) = std::str::from_utf8(&mutated) {
            drop(chart_read(xml));
        }
    }
    let parsed = chart_read(CHART_SEED)?;
    assert_eq!(parsed.kind(), chart::Kind::Chart);
    assert_eq!(parsed.all_text(), "Revenue");
    Ok(())
}

#[test]
fn chart_malformed_inputs_yield_typed_errors() {
    assert_invalid(
        chart_read(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart><chart:plot-area></chart:chart></chart:plot-area></office:chart></office:body></office:document-content>"#,
        ),
        "mismatched chart close",
    );
    assert_invalid(
        chart_read(
            r#"<!DOCTYPE office [<!ENTITY x "y">]><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><office:chart><chart:chart/></office:chart></office:body></office:document-content>"#,
        ),
        "chart DTD",
    );
    assert_invalid(
        chart_read(
            r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"><office:body><chart:chart/></office:body></office:document-content>"#,
        ),
        "misplaced chart",
    );
}

#[test]
fn drawing_resource_parsers_report_the_required_raw_xml_policy() {
    let xml = "<draw:gradient/><draw:hatch/><draw:marker/><draw:stroke-dash/><draw:opacity/><draw:fill-image/>"
        .repeat(2);
    let limits =
        XmlStreamLimits::new(64, 256, 4_000, 1_000, 64 * 1024).expect("test policy is finite");
    let expected = |error: OdfError, input_len: usize| {
        assert!(matches!(
            error,
            OdfError::Limit {
                kind: litchi_odf_common::LimitKind::RawXml(
                    litchi_odf_common::RawXmlLimitKind::InputBytes
                ),
                observed,
                maximum,
            } if observed == input_len as u64 && maximum == 64
        ));
    };
    expected(
        drawing::resources::gradient::parse_drawing_gradients(&xml, limits)
            .expect_err("gradient input must be bounded"),
        xml.len(),
    );
    expected(
        drawing::resources::hatch::parse_drawing_hatches(&xml, limits)
            .expect_err("hatch input must be bounded"),
        xml.len(),
    );
    expected(
        drawing::resources::marker::parse_drawing_markers(&xml, limits)
            .expect_err("marker input must be bounded"),
        xml.len(),
    );
    expected(
        drawing::resources::stroke_dash::parse_drawing_stroke_dashes(&xml, limits)
            .expect_err("stroke-dash input must be bounded"),
        xml.len(),
    );
    expected(
        drawing::resources::opacity::parse_drawing_opacities(&xml, limits)
            .expect_err("opacity input must be bounded"),
        xml.len(),
    );
    expected(
        drawing::resources::fill_image::parse_drawing_fill_images(&xml, limits)
            .expect_err("fill-image input must be bounded"),
        xml.len(),
    );

    let unrelated = "x".repeat(65);
    expected(
        drawing::resources::gradient::parse_drawing_gradients(&unrelated, limits)
            .expect_err("unrelated input must still be bounded"),
        unrelated.len(),
    );
}

#[test]
fn metadata_and_manifest_report_the_required_raw_xml_policy() {
    let metadata_xml = r#"<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#;
    let manifest_xml = r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/"/></m:manifest>"#;
    let limits =
        XmlStreamLimits::new(64, 256, 4_000, 1_000, 64 * 1024).expect("test policy is finite");
    let expected = |error: OdfError, input_len: usize| {
        assert!(matches!(
            error,
            OdfError::Limit {
                kind: litchi_odf_common::LimitKind::RawXml(
                    litchi_odf_common::RawXmlLimitKind::InputBytes
                ),
                observed,
                maximum,
            } if observed == input_len as u64 && maximum == 64
        ));
    };
    expected(
        OdfError::from(
            OdfMetadata::from_xml(metadata_xml, limits)
                .expect_err("metadata input must be bounded"),
        ),
        metadata_xml.len(),
    );
    expected(
        OdfError::from(
            parse_manifest(manifest_xml, limits).expect_err("manifest input must be bounded"),
        ),
        manifest_xml.len(),
    );
}

#[test]
fn shared_media_and_object_scanners_are_hardened() -> Result<(), Error> {
    let images = media::scan_flat(FLAT_MEDIA_SEED)?;
    let objects = embedded::scan_flat(FLAT_MEDIA_SEED)?;
    assert_eq!(images.len(), 1);
    assert_eq!(objects.len(), 1);

    for end in 0..FLAT_MEDIA_SEED.len() {
        drop(media::scan_flat(&FLAT_MEDIA_SEED[..end]));
        drop(embedded::scan_flat(&FLAT_MEDIA_SEED[..end]));
    }
    for position in 0..FLAT_MEDIA_SEED.len() {
        let mut mutated = FLAT_MEDIA_SEED.as_bytes().to_vec();
        mutated[position] ^= 1;
        if let Ok(xml) = std::str::from_utf8(&mutated) {
            drop(media::scan_flat(xml));
            drop(embedded::scan_flat(xml));
        }
    }

    let dtd = r#"<!DOCTYPE office [<!ENTITY x "y">]><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text>&x;</office:text></office:body></office:document>"#;
    assert_invalid(media::scan_flat(dtd), "image scanner DTD");
    assert_invalid(embedded::scan_flat(dtd), "object scanner DTD");
    Ok(())
}

#[test]
fn archive_and_manifest_malformed_sweeps_never_panic() -> Result<(), Box<dyn std::error::Error>> {
    let bytes = package()?;
    for end in 0..bytes.len() {
        exercise_archive(&bytes[..end]);
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.clone();
        mutated[position] ^= 1;
        exercise_archive(&mutated);
    }
    exercise_archive(&bytes);

    assert_invalid(
        parse_manifest(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="x"/><m:file-entry m:full-path="x"/></m:manifest>"#,
            XmlStreamLimits::default(),
        ),
        "duplicate manifest path",
    );
    assert_invalid(
        parse_manifest(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="x" m:size="invalid"/></m:manifest>"#,
            XmlStreamLimits::default(),
        ),
        "invalid manifest size",
    );
    Ok(())
}

fn chart_read(xml: &str) -> Result<chart::Element, Error> {
    chart::read(xml, XmlStreamLimits::default()).map_err(Error::from)
}
