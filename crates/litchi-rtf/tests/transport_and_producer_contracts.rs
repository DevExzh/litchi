#![allow(
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "fixed fixtures and test assertions are intentionally direct"
)]

use litchi_rtf::{
    Document,
    edit::Error,
    transport::{compress, decompress, is_compressed_rtf},
};

const LZFU_RAW: &[u8] = b"{\\rtf1\\ansi\\ansicpg1252\\pard hello world}\r\n";
const LZFU_SPEC: &[u8] = &[
    0x2d, 0x00, 0x00, 0x00, 0x2b, 0x00, 0x00, 0x00, 0x4c, 0x5a, 0x46, 0x75, 0xf1, 0xc5, 0xc7, 0xa7,
    0x03, 0x00, 0x0a, 0x00, 0x72, 0x63, 0x70, 0x67, 0x31, 0x32, 0x35, 0x42, 0x32, 0x0a, 0xf3, 0x20,
    0x68, 0x65, 0x6c, 0x09, 0x00, 0x20, 0x62, 0x77, 0x05, 0xb0, 0x6c, 0x64, 0x7d, 0x0a, 0x80, 0x0f,
    0xa0,
];
const NATIVE_SENTINEL: &str = "Litchi native resave 2026-08-10";

#[test]
fn literal_byte1252_transport_is_exact_and_changed_write_is_atomic_on_refusal() {
    let mut source = br"{\rtf1\ansi\ansicpg1252 caf".to_vec();
    source.push(0xe9);
    source.push(b'}');

    let document = Document::from_bytes(&source).unwrap();
    assert_eq!(document.text(), "caf\u{e9}");
    assert_eq!(document.body().len(), document.text().len());
    assert_eq!(document.to_bytes().unwrap(), source);

    let noop = document.edit().commit().unwrap();
    assert!(noop.snapshot().same_snapshot(&document));
    assert_eq!(noop.snapshot().to_bytes().unwrap(), source);

    let mut edit = document.edit();
    edit.replace_paragraph_text(0, "changed").unwrap();
    assert!(matches!(edit.commit(), Err(Error::Rtf(_))));
    assert_eq!(document.to_bytes().unwrap(), source);
}

#[test]
fn specification_lzfu_transport_is_exact_and_changed_write_fails_closed() {
    assert!(is_compressed_rtf(LZFU_SPEC));
    assert_eq!(decompress(LZFU_SPEC).unwrap(), LZFU_RAW);
    assert_eq!(compress(LZFU_RAW, true).unwrap(), LZFU_SPEC);

    let document = Document::from_bytes(LZFU_SPEC).unwrap();
    assert_eq!(document.text(), "hello world");
    assert_eq!(document.body().len(), document.text().len());
    assert_eq!(document.to_bytes().unwrap(), LZFU_SPEC);
    let mut streamed = Vec::new();
    document.write_to(&mut streamed).unwrap();
    assert_eq!(streamed, LZFU_SPEC);

    let noop = document.edit().commit().unwrap();
    assert!(noop.snapshot().same_snapshot(&document));
    assert_eq!(noop.snapshot().to_bytes().unwrap(), LZFU_SPEC);

    let mut edit = document.edit();
    edit.replace_paragraph_text(0, "changed").unwrap();
    assert!(matches!(edit.commit(), Err(Error::UnsupportedSource(_))));
    assert_eq!(document.to_bytes().unwrap(), LZFU_SPEC);
}

#[test]
fn libreoffice_watermark_is_visible_through_public_header_shapes_and_exact_on_noop() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/watermark.rtf"
    );
    let document = Document::from_bytes(source).unwrap();
    assert_eq!(document.body().len(), document.text().len());
    let themed_fill = document
        .sections()
        .iter()
        .flat_map(|section| &section.headers_footers)
        .flat_map(|header_footer| &header_footer.shapes)
        .flat_map(|shape| &shape.properties)
        .find(|property| property.name == "fillColor" && property.theme_value.is_some())
        .unwrap();
    assert_eq!(themed_fill.value, "4626167");
    assert_eq!(document.to_bytes().unwrap(), source);

    let noop = document.edit().commit().unwrap();
    assert!(noop.snapshot().same_snapshot(&document));
    let mut streamed = Vec::new();
    noop.snapshot().write_to(&mut streamed).unwrap();
    assert_eq!(streamed, source);
    let reopened = Document::from_bytes(&streamed).unwrap();
    assert_eq!(reopened.to_bytes().unwrap(), source);
}

#[test]
fn relsize_public_edit_and_checked_native_resave_reopen_semantically() {
    let source_bytes = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf"
    );
    let changed_bytes =
        include_bytes!("../../../test-data/office-interop/litchi-changed/relsize-litchi.rtf");
    let resaved_bytes =
        include_bytes!("../../../test-data/office-interop/libreoffice-resaved/relsize-litchi.rtf");

    let source = Document::from_bytes(source_bytes).unwrap();
    assert_eq!(source.shapes().len(), 1);
    assert_eq!(source.shapes()[0].text, "Textbox text.\n");
    assert_eq!(source.to_bytes().unwrap(), source_bytes);
    let original = source.shapes()[0].clone();

    let mut edit = source.edit();
    edit.set_shape_text(0, NATIVE_SENTINEL).unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().operation_count(), 1);
    assert!(commit.diagnostics().changed());
    let generated = commit.snapshot().to_bytes().unwrap();
    let generated_reopen = Document::from_bytes(&generated).unwrap();
    let generated_shape = &generated_reopen.shapes()[0];
    assert_eq!(generated_shape.text, NATIVE_SENTINEL);
    assert_eq!(generated_shape.position, original.position);
    assert_eq!(generated_shape.geometry, original.geometry);
    assert_eq!(generated_shape.properties, original.properties);
    assert_eq!(generated_shape.text_formatting, original.text_formatting);
    assert_eq!(generated_reopen.text(), source.text());
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let changed = Document::from_bytes(changed_bytes).unwrap();
    assert_eq!(changed.shapes().len(), 1);
    assert_eq!(changed.shapes()[0].text, NATIVE_SENTINEL);
    assert_eq!(changed.to_bytes().unwrap(), changed_bytes);

    let resaved = Document::from_bytes(resaved_bytes).unwrap();
    assert_eq!(resaved.shapes().len(), 1);
    assert_eq!(resaved.shapes()[0].text, format!("{NATIVE_SENTINEL}\n"));
    assert_eq!(resaved.to_bytes().unwrap(), resaved_bytes);
}
