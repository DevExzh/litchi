use litchi_odf::{
    Document, DocumentBuilder, FlatOpenDocument, MutableDocument, OpenDocumentPackage, StyleColumn,
    StyleColumnLength, StyleColumnSeparator, StyleColumnSeparatorAlignment,
    StyleColumnSeparatorStyle, StyleColumns, parse_style_columns,
};
use std::io::{Cursor, Write};
use zip::CompressionMethod;
use zip::write::SimpleFileOptions;

const STYLE: &str = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const FO: &str = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";

fn wrap(body: &str) -> String {
    format!(
        r#"<s:section-properties xmlns:s="{STYLE}" xmlns:f="{FO}">{body}</s:section-properties>"#
    )
}

#[test]
fn parses_libreoffice_section_separator_and_other_property_contexts() {
    let section = include_str!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/uitest/data/section-columns-separator.fodt"
    ));
    let parsed = parse_style_columns(section).unwrap();
    let columns = &parsed[0];
    assert_eq!(columns.column_count, 2);
    assert_eq!(columns.column_gap.as_ref().unwrap().as_str(), "0.6cm");
    assert_eq!(columns.columns()[0].relative_width, 32_767);
    let separator = columns.separator.as_ref().unwrap();
    assert_eq!(separator.width.as_str(), "0.009cm");
    assert_eq!(separator.style, Some(StyleColumnSeparatorStyle::Dotted));
    assert_eq!(separator.height_percent, Some(50));
    assert_eq!(
        separator.vertical_alignment,
        Some(StyleColumnSeparatorAlignment::Bottom)
    );
    assert_eq!(separator.color, Some((0x99, 0xAA, 0xBB)));

    for fixture in [
        include_str!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/pageColumns.fodt"
        )),
        include_str!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/libreoffice-core/sw/qa/extras/pagelinespacing/data/frameWithColumns.fodt"
        )),
    ] {
        let flat = FlatOpenDocument::from_bytes(fixture.as_bytes().to_vec()).unwrap();
        let parsed = flat.style_columns().unwrap();
        assert!(
            parsed
                .iter()
                .any(|columns| columns.column_count == 2 && columns.columns().len() == 2)
        );
    }
}

#[test]
fn parses_aliases_equal_width_and_round_trips_deterministically() {
    let equal = parse_style_columns(&wrap(r#"<s:columns f:column-count="3"/>"#)).unwrap();
    assert_eq!(equal.len(), 1);
    assert!(equal[0].columns().is_empty());

    let xml = wrap(
        r##"<s:columns f:column-count="2" f:column-gap="0.5cm"><s:column-sep s:width="0.01cm" s:style="dot-dashed" s:height="75%" s:vertical-align="top" s:color="#abcdef"/><s:column s:rel-width="1*" f:start-indent="0cm" f:end-indent=".2cm"/><s:column s:rel-width="2*" f:space-before="0.1cm" f:space-after="0.2cm"/></s:columns>"##,
    );
    let parsed = parse_style_columns(&xml).unwrap().remove(0);
    let fragment = parsed.to_xml_fragment().unwrap();
    assert!(fragment.starts_with(r#"<style:columns xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" fo:column-count="2" fo:column-gap="0.5cm">"#));
    assert!(fragment.contains(r##"<style:column-sep style:color="#ABCDEF" style:height="75%" style:style="dot-dashed" style:vertical-align="top" style:width="0.01cm"/>"##));
    assert_eq!(parse_style_columns(&wrap(&fragment)).unwrap(), vec![parsed]);
}

#[test]
fn rejects_wrong_namespace_order_cardinality_values_and_caps() {
    let malformed = [
        format!(r#"<s:columns xmlns:s="{STYLE}" xmlns:f="{FO}" f:column-count="2"/>"#),
        wrap(r#"<s:columns/>"#),
        wrap(r#"<s:columns f:column-count="0"/>"#),
        wrap(r#"<s:columns f:column-count="65"/>"#),
        wrap(r#"<s:columns f:column-count="2"><s:column s:rel-width="1*"/></s:columns>"#),
        wrap(r#"<s:columns f:column-count="1"><s:column s:rel-width="0*"/></s:columns>"#),
        wrap(r#"<s:columns f:column-count="1"><s:column s:rel-width="1cm"/></s:columns>"#),
        wrap(
            r#"<s:columns f:column-count="1"><s:column s:rel-width="1*"/><s:column-sep s:width="1pt"/></s:columns>"#,
        ),
        wrap(
            r#"<s:columns f:column-count="1"><s:column-sep s:width="1pt"/><s:column-sep s:width="1pt"/><s:column s:rel-width="1*"/></s:columns>"#,
        ),
        wrap(
            r#"<s:columns f:column-count="1"><s:column-sep/><s:column s:rel-width="1*"/></s:columns>"#,
        ),
        wrap(
            r#"<s:columns f:column-count="1"><s:column-sep s:width="1pt" s:height="101%"/><s:column s:rel-width="1*"/></s:columns>"#,
        ),
        wrap(
            r#"<s:columns f:column-count="1" s:unknown="x"><s:column s:rel-width="1*"/></s:columns>"#,
        ),
        format!(
            r#"<s:section-properties xmlns:s="{STYLE}" xmlns:f="{FO}" xmlns:x="urn:wrong"><x:columns f:column-count="1"/></s:section-properties>"#
        ),
        format!(
            r#"<!DOCTYPE x><s:section-properties xmlns:s="{STYLE}" xmlns:f="{FO}"><s:columns f:column-count="1"/></s:section-properties>"#
        ),
    ];
    for xml in malformed {
        assert!(parse_style_columns(&xml).is_err(), "accepted {xml}");
    }

    let explicit = (0..65)
        .map(|_| r#"<s:column s:rel-width="1*"/>"#)
        .collect::<String>();
    let overflow = wrap(&format!(
        r#"<s:columns f:column-count="64">{explicit}</s:columns>"#
    ));
    assert!(parse_style_columns(&overflow).is_err());
}

#[test]
fn builder_package_page_layout_and_mutable_update_round_trip() {
    let mut left = StyleColumn::new(1).unwrap();
    left.end_indent = Some(StyleColumnLength::new("0.2cm").unwrap());
    let mut right = StyleColumn::new(1).unwrap();
    right.start_indent = Some(StyleColumnLength::new("0.2cm").unwrap());
    let mut columns = StyleColumns::try_with_columns(2, vec![left, right]).unwrap();
    columns.column_gap = Some(StyleColumnLength::new("0.4cm").unwrap());
    let mut separator = StyleColumnSeparator::new(StyleColumnLength::new("0.01cm").unwrap());
    separator.style = Some(StyleColumnSeparatorStyle::Solid);
    columns.separator = Some(separator);

    let mut builder = DocumentBuilder::new();
    builder
        .add_page_layout_columns("Columns&Layout", columns.clone())
        .unwrap();
    builder.add_paragraph("body").unwrap();
    let bytes = builder.build().unwrap();
    let package = OpenDocumentPackage::from_bytes(bytes.clone()).unwrap();
    assert!(package.style_columns().unwrap().contains(&columns));
    let document = Document::from_bytes(bytes).unwrap();
    let layout = document
        .page_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Columns&Layout")
        .unwrap();
    assert_eq!(layout.properties.unwrap().columns, Some(columns));

    let mut mutable = MutableDocument::from_document(document).unwrap();
    let replacement = StyleColumns::new(3).unwrap();
    mutable
        .set_page_layout_columns("Columns&Layout", &replacement)
        .unwrap();
    let output = mutable.to_bytes().unwrap();
    let document = Document::from_bytes(output).unwrap();
    let layout = document
        .page_layouts()
        .unwrap()
        .into_iter()
        .find(|layout| layout.name == "Columns&Layout")
        .unwrap();
    assert_eq!(layout.properties.unwrap().columns, Some(replacement));
}

#[test]
fn mutable_update_preserves_inherited_aliases() {
    let styles = format!(
        r#"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="{STYLE}" xmlns:f="{FO}" o:version="1.3"><o:automatic-styles><s:page-layout s:name="Alias"><s:page-layout-properties><s:columns f:column-count="2"/></s:page-layout-properties></s:page-layout></o:automatic-styles><o:master-styles/></o:document-styles>"#
    );
    let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:version="1.3"><o:body><o:text><t:p>body</t:p></o:text></o:body></o:document-content>"#;
    let mut archive = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    archive.start_file("mimetype", stored).unwrap();
    archive
        .write_all(b"application/vnd.oasis.opendocument.text")
        .unwrap();
    archive.start_file("content.xml", deflated).unwrap();
    archive.write_all(content.as_bytes()).unwrap();
    archive.start_file("styles.xml", deflated).unwrap();
    archive.write_all(styles.as_bytes()).unwrap();
    archive
        .start_file("META-INF/manifest.xml", deflated)
        .unwrap();
    archive.write_all(br#"<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/></manifest:manifest>"#).unwrap();
    let document = Document::from_bytes(archive.finish().unwrap().into_inner()).unwrap();
    assert_eq!(
        document.page_layouts().unwrap()[0]
            .properties
            .as_ref()
            .unwrap()
            .columns
            .as_ref()
            .unwrap()
            .column_count,
        2
    );
    let mut mutable = MutableDocument::from_document(document).unwrap();
    mutable
        .set_page_layout_columns("Alias", &StyleColumns::new(4).unwrap())
        .unwrap();
    let round_trip = Document::from_bytes(mutable.to_bytes().unwrap()).unwrap();
    assert_eq!(
        round_trip.page_layouts().unwrap()[0]
            .properties
            .as_ref()
            .unwrap()
            .columns
            .as_ref()
            .unwrap()
            .column_count,
        4
    );
}
