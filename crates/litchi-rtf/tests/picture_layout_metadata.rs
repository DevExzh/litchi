#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{ImageType, Picture, RtfDocument, RtfWriter};
use std::borrow::Cow;

fn write(document: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(document).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn bitmap_crop_and_source_header_metadata_round_trip_inertly() {
    let source = concat!(
        r#"{\rtf1{\*\shppict{\pict\wbitmap0\picw2\pich2"#,
        r#"\picwgoal100\pichgoal120\picscalex125\picscaley80\picscaled"#,
        r#"\piccropl-1\piccropr2\piccropt3\piccropb-4"#,
        r#"\picbmp\picbpp24\wbmbitspixel8\wbmplanes1\wbmwidthbytes8 00}}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let picture = &document.pictures()[0];
    assert_eq!(picture.image_type, ImageType::Dib);
    assert!(picture.scaled);
    assert_eq!(picture.crop.left, Some(-1));
    assert_eq!(picture.crop.right, Some(2));
    assert_eq!(picture.crop.top, Some(3));
    assert_eq!(picture.crop.bottom, Some(-4));
    assert!(picture.bitmap.windows_bitmap);
    assert!(picture.bitmap.bitmap_source);
    assert_eq!(picture.bitmap.bits_per_pixel, Some(24));
    assert_eq!(picture.bitmap.windows_bits_per_pixel, Some(8));
    assert_eq!(picture.bitmap.planes, Some(1));
    assert_eq!(picture.bitmap.width_bytes, Some(8));

    let serialized = write(&document);
    assert!(serialized.contains("\\wbitmap0"));
    assert!(serialized.contains("\\picscaled"));
    assert!(serialized.contains("\\piccropl-1"));
    assert!(serialized.contains("\\wbmwidthbytes8"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.pictures(), document.pictures());
    assert_eq!(
        reparsed.picture_compatibility_records(),
        document.picture_compatibility_records()
    );
}

#[test]
fn picture_layout_controls_reject_ambiguous_or_unbounded_forms() {
    for source in [
        r"{\rtf1{\pict\wbitmap0\piccropl 00}}",
        r"{\rtf1{\pict\wbitmap0\picbpp 00}}",
        r"{\rtf1{\pict\wbitmap0\wbmplanes0 00}}",
        r"{\rtf1{\pict\wbitmap0\wbmwidthbytes-1 00}}",
        r"{\rtf1{\pict\wbitmap0\picscaled1 00}}",
        r"{\rtf1{\pict\wbitmap0\picbmp1 00}}",
        r"{\rtf1{\pict\wbitmap0\piccropl1\piccropl2 00}}",
        r"{\rtf1{\pict\wbitmap0 00\piccropb1}}",
        r"{\rtf1{\pict\pngblip\picbpp8 00}}",
        r"{\rtf1\piccropl1 Body}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted hostile source: {source}"
        );
    }

    let mut picture = Picture::new(ImageType::Png, Cow::Borrowed(&[1]));
    picture.bitmap.bits_per_pixel = Some(8);
    assert!(picture.validate().is_err());

    let mut picture = Picture::new(ImageType::Dib, Cow::Borrowed(&[1]));
    picture.bitmap.width_bytes = Some(i32::MAX as u32 + 1);
    assert!(picture.validate().is_err());
}
