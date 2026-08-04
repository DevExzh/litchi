//! Fuzz-style hardening for the ODP, ODG, ODC, and ODM parsers.
//!
//! Same contract as `parser_hardening.rs`: truncated, mutated, and
//! hand-crafted malformed documents must always produce a typed `Result` —
//! `Ok` or `Error::InvalidFormat` — and never a panic.

use litchi_odf::{FlatChartDocument, FlatDrawingDocument, FlatPresentation, MasterDocument};

/// Feature-rich flat ODP seed: declarations, a slide with a transition, an
/// animation tree, grouped and custom shapes, a media plugin, and notes.
const ODP_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:anim="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.presentation" office:version="1.3"><office:automatic-styles><style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="automatic" smil:type="fade" smil:subtype="crossfade" presentation:duration="PT2S"/></style:style></office:automatic-styles><office:body><office:presentation><presentation:declarations><presentation:placeholder-declaration presentation:class="title" draw:name="title1"/></presentation:declarations><draw:page draw:name="s1" draw:style-name="dp1" draw:master-page-name="Standard"><draw:frame draw:name="media" draw:layer="layout" svg:x="1cm" svg:y="1cm" svg:width="4cm" svg:height="3cm"><draw:plugin xlink:href="movie.avi" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad" draw:mime-type="video/avi"/></draw:frame><draw:g draw:name="group1" draw:id="shape1"><draw:custom-shape draw:name="cs" svg:x="2cm" svg:y="2cm" svg:width="3cm" svg:height="3cm"><draw:enhanced-geometry draw:mirror-horizontal="false" draw:type="rectangle"><draw:equation draw:name="f1" draw:formula="width/2"/><draw:handle draw:handle-position="$0 0"/></draw:enhanced-geometry><text:p>custom</text:p></draw:custom-shape></draw:g><presentation:notes><text:p>speaker notes</text:p></presentation:notes><anim:par><anim:seq><anim:par><anim:animate smil:attributeName="visibility" smil:to="visible" smil:dur="0.5s" smil:targetElement="shape1"/></anim:par></anim:seq></anim:par></draw:page></office:presentation></office:body></office:document>"#;

/// Feature-rich flat ODG seed: layers, nested groups, and a custom shape
/// with enhanced geometry.
const ODG_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics" office:version="1.3"><office:master-styles><draw:layer-set><draw:layer draw:name="layout"/><draw:layer draw:name="background"/></draw:layer-set></office:master-styles><office:body><office:drawing><draw:page draw:name="p1"><draw:g draw:name="outer"><draw:rect draw:layer="layout" svg:x="1cm" svg:y="1cm" svg:width="2cm" svg:height="2cm"/><draw:g draw:name="inner"><draw:custom-shape svg:x="3cm" svg:y="3cm" svg:width="4cm" svg:height="4cm"><draw:enhanced-geometry draw:type="round-rectangle" draw:modifiers="30"><draw:equation draw:name="adj" draw:formula="min(width,height)*?0/100"/><draw:handle draw:handle-position="$0 top"/></draw:enhanced-geometry><text:p>rounded</text:p></draw:custom-shape></draw:g></draw:g><draw:ellipse draw:layer="background" svg:cx="5cm" svg:cy="5cm" svg:rx="2cm" svg:ry="1cm"/></draw:page></office:drawing></office:body></office:document>"#;

/// Feature-rich flat ODC seed: title, legend, axes, a series with data
/// points, and a cached table.
const ODC_SEED: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:bar"><chart:title><text:p>Revenue</text:p></chart:title><chart:legend chart:legend-position="end"/><chart:plot-area><chart:axis chart:dimension="x" chart:name="primary-x"/><chart:axis chart:dimension="y" chart:name="primary-y"/><chart:series chart:values-cell-range-address="Sheet1.$B$1:.$B$3" chart:label-cell-address="Sheet1.$B$4" chart:domain-cell-range-address="Sheet1.$A$1:.$A$3"><chart:data-point chart:repeated="2"/></chart:series><table:table table:name="local-table"><table:table-row><table:table-cell office:value-type="float" office:value="4"><text:p>4</text:p></table:table-cell></table:table-row></table:table></chart:plot-area></chart:chart></office:chart></office:body></office:document>"#;

/// ODM content.xml for the packaged master-document seed.
const ODM_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.3"><office:body><office:text><text:section text:name="intro" xml:id="sec_intro"><text:section-source xlink:href="intro.odt" xlink:type="simple" text:section-name="chapter1"/><text:p>placeholder</text:p></text:section><text:p>after</text:p></office:text></office:body></office:document-content>"#;

fn package(mimetype: &str, content: &str) -> Vec<u8> {
    let mut output = std::io::Cursor::new(Vec::new());
    let mut zip = zip::ZipWriter::new(&mut output);
    let stored =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    let deflated = zip::write::SimpleFileOptions::default()
        .compression_method(zip::CompressionMethod::Deflated);
    zip.start_file("mimetype", stored).unwrap();
    std::io::Write::write_all(&mut zip, mimetype.as_bytes()).unwrap();
    zip.start_file("content.xml", deflated).unwrap();
    std::io::Write::write_all(&mut zip, content.as_bytes()).unwrap();
    zip.start_file("META-INF/manifest.xml", deflated).unwrap();
    std::io::Write::write_all(
        &mut zip,
        format!(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" m:version="1.3"><m:file-entry m:full-path="/" m:media-type="{mimetype}"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml"/></m:manifest>"#
        )
        .as_bytes(),
    )
    .unwrap();
    zip.finish().unwrap();
    output.into_inner()
}

fn exercise_odp(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    let mut flat = match FlatPresentation::from_bytes(bytes) {
        Ok(flat) => flat,
        Err(error) => return Some(error),
    };
    let presentation = flat.presentation_mut();
    for result in [
        presentation.slides().map(|_| ()),
        presentation.declarations().map(|_| ()),
        presentation.page_layouts().map(|_| ()),
        presentation.forms().map(|_| ()),
        presentation.images().map(|_| ()),
    ] {
        if let Err(error) = result {
            return Some(error);
        }
    }
    None
}

fn exercise_odg(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    match FlatDrawingDocument::from_bytes(bytes) {
        Ok(flat) => {
            let _ = flat.drawing().pages();
            let _ = flat.metadata();
            None
        },
        Err(error) => Some(error),
    }
}

fn exercise_odc(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    let mut flat = match FlatChartDocument::from_bytes(bytes) {
        Ok(flat) => flat,
        Err(error) => return Some(error),
    };
    let chart = flat.chart_mut();
    if let Some(plot_area) = chart.plot_area() {
        for axis in plot_area.axes() {
            if let Err(error) = axis.dimension() {
                return Some(error);
            }
        }
        for series in plot_area.series() {
            let _ = series.values_range();
            let _ = series.attached_axis();
        }
    }
    let _ = chart.legend();
    None
}

fn exercise_odm(bytes: Vec<u8>) -> Option<litchi_core::Error> {
    match MasterDocument::from_bytes(bytes) {
        Ok(master) => {
            let _ = master.subdocuments();
            let _ = master.global();
            if let Err(error) = master.document().text() {
                return Some(error);
            }
            None
        },
        Err(error) => Some(error),
    }
}

fn assert_typed_error(error: Option<litchi_core::Error>, case: &str) {
    let Some(error) = error else {
        panic!("case '{case}' unexpectedly parsed");
    };
    assert!(
        matches!(error, litchi_core::Error::InvalidFormat(_)),
        "case '{case}' produced a non-typed error: {error:?}"
    );
}

#[test]
fn odp_truncation_and_mutation_sweeps_never_panic() {
    let bytes = ODP_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise_odp(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_odp(mutated);
    }
    assert!(exercise_odp(bytes.to_vec()).is_none());
}

#[test]
fn odg_truncation_and_mutation_sweeps_never_panic() {
    let bytes = ODG_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise_odg(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_odg(mutated);
    }
    assert!(exercise_odg(bytes.to_vec()).is_none());
}

#[test]
fn odc_truncation_and_mutation_sweeps_never_panic() {
    let bytes = ODC_SEED.as_bytes();
    for end in 0..bytes.len() {
        exercise_odc(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_odc(mutated);
    }
    assert!(exercise_odc(bytes.to_vec()).is_none());
}

#[test]
fn odm_truncation_and_mutation_sweeps_never_panic() {
    let bytes = package(
        "application/vnd.oasis.opendocument.text-master",
        ODM_CONTENT,
    );
    // Package sweeps corrupt the zip structure; content sweeps corrupt the
    // XML inside a valid package.
    for end in 0..bytes.len() {
        exercise_odm(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 0x01;
        exercise_odm(mutated);
    }
    assert!(exercise_odm(bytes).is_none());
}

#[test]
fn odp_malformed_inputs_yield_typed_errors() {
    let fodp = |automatic_styles: &str, body: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:smil="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" office:mimetype="application/vnd.oasis.opendocument.presentation" office:version="1.3"><office:automatic-styles>{automatic_styles}</office:automatic-styles><office:body><office:presentation>{body}</office:presentation></office:body></office:document>"#
        )
        .into_bytes()
    };
    // Unknown presentation:transition-type in drawing-page style properties.
    assert_typed_error(
        exercise_odp(fodp(
            r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="not-a-type"/></style:style>"#,
            r#"<draw:page draw:name="s" draw:style-name="dp1"><draw:rect/></draw:page>"#,
        )),
        "invalid transition type",
    );
    // Unknown presentation:transition-speed in drawing-page style properties.
    assert_typed_error(
        exercise_odp(fodp(
            r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-speed="warp"/></style:style>"#,
            r#"<draw:page draw:name="s" draw:style-name="dp1"><draw:rect/></draw:page>"#,
        )),
        "invalid transition speed",
    );
    // A transition sound without its required xlink:href.
    assert_typed_error(
        exercise_odp(fodp(
            r#"<style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties presentation:transition-type="manual"><presentation:sound/></style:drawing-page-properties></style:style>"#,
            r#"<draw:page draw:name="s" draw:style-name="dp1"><draw:rect/></draw:page>"#,
        )),
        "transition sound without href",
    );
    // Mismatched closing tag inside a slide.
    assert_typed_error(
        exercise_odp(fodp(
            "",
            r#"<draw:page draw:name="s"><draw:g><draw:rect/></draw:page></draw:g>"#,
        )),
        "mismatched slide close",
    );
}

#[test]
fn odg_malformed_inputs_yield_typed_errors() {
    let fodg = |body: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics" office:version="1.3"><office:body>{body}</office:body></office:document>"#
        )
        .into_bytes()
    };
    // A page outside office:drawing.
    assert_typed_error(
        exercise_odg(fodg(r#"<draw:page draw:name="s"/>"#)),
        "page outside office:drawing",
    );
    // Two drawing roots.
    assert_typed_error(
        exercise_odg(fodg(
            r#"<office:drawing><draw:page draw:name="a"/></office:drawing><office:drawing><draw:page draw:name="b"/></office:drawing>"#,
        )),
        "duplicate drawing roots",
    );
    // An empty document element cannot be the drawing root.
    assert_typed_error(
        exercise_odg(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.graphics" office:version="1.3"/>"#
                .as_bytes()
                .to_vec(),
        ),
        "empty drawing root",
    );
}

#[test]
fn odc_malformed_inputs_yield_typed_errors() {
    let fodc = |inner: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3"><office:body>{inner}</office:body></office:document-content>"#
        )
        .into_bytes()
    };
    // A root that is not office:document-content.
    assert_typed_error(
        exercise_odc(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3"><office:body><office:chart><chart:chart/></office:chart></office:body></office:document>"#
                .as_bytes()
                .to_vec(),
        ),
        "wrong chart root",
    );
    // Duplicate office:chart wrappers.
    assert_typed_error(
        exercise_odc(fodc(
            r#"<office:chart><chart:chart chart:class="chart:bar"/></office:chart><office:chart><chart:chart chart:class="chart:line"/></office:chart>"#,
        )),
        "duplicate office:chart",
    );
    // An axis with an unknown dimension is a typed semantic error.
    assert_typed_error(
        exercise_odc(fodc(
            r#"<office:chart><chart:chart chart:class="chart:bar"><chart:plot-area><chart:axis chart:dimension="z" chart:name="bad"/></chart:plot-area></chart:chart></office:chart>"#,
        )),
        "invalid axis dimension",
    );
}

#[test]
fn odm_malformed_inputs_yield_typed_errors() {
    const MIME: &str = "application/vnd.oasis.opendocument.text-master";
    let content = |inner: &str| {
        format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3"><office:body><office:text>{inner}</office:text></office:body></office:document-content>"#
        )
    };
    // Duplicate master section names.
    assert_typed_error(
        exercise_odm(package(
            MIME,
            &content(
                r#"<text:section text:name="dup"><text:p>a</text:p></text:section><text:section text:name="dup"><text:p>b</text:p></text:section>"#,
            ),
        )),
        "duplicate section names",
    );
    // Duplicate xml:id values.
    assert_typed_error(
        exercise_odm(package(
            MIME,
            &content(
                r#"<text:section text:name="a" xml:id="same"><text:p>a</text:p></text:section><text:section text:name="b" xml:id="same"><text:p>b</text:p></text:section>"#,
            ),
        )),
        "duplicate xml:id",
    );
    // DOCTYPE is rejected outright.
    assert_typed_error(
        exercise_odm(package(
            MIME,
            r#"<?xml version="1.0"?><!DOCTYPE office><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3"><office:body><office:text/></office:body></office:document-content>"#,
        )),
        "DOCTYPE in master content",
    );
    // A wrong mimetype is not a master document.
    assert_typed_error(
        exercise_odm(package(
            "application/vnd.oasis.opendocument.text",
            &content("<text:p>plain</text:p>"),
        )),
        "wrong mimetype",
    );
}

#[test]
fn misplaced_but_wellformed_inputs_do_not_panic() {
    // Foreign elements inside a chart subtree are preserved, not panicked on.
    exercise_odc(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0" office:mimetype="application/vnd.oasis.opendocument.chart" office:version="1.3"><office:body><office:chart><chart:chart chart:class="chart:bar"><ext:thing xmlns:ext="urn:example" ext:attr="1"/></chart:chart></office:chart></office:body></office:document-content>"#
            .as_bytes()
            .to_vec(),
    );
    // Invalid UTF-8 is a typed error, never a panic.
    let mut bytes = ODP_SEED.as_bytes().to_vec();
    bytes.insert(bytes.len() / 2, 0xFF);
    exercise_odp(bytes);
    let mut bytes = ODG_SEED.as_bytes().to_vec();
    bytes.insert(bytes.len() / 2, 0xFE);
    exercise_odg(bytes);
    let mut bytes = ODC_SEED.as_bytes().to_vec();
    bytes.insert(bytes.len() / 2, 0x80);
    exercise_odc(bytes);
    // Corrupt zip containers are typed errors.
    exercise_odm(b"not a zip at all".to_vec());
    exercise_odm(Vec::new());
}
