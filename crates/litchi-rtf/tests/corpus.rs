#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;

const FIXTURES: &[&str] = &[
    "background.rtf",
    "chtoutline.rtf",
    "cjklist12.rtf",
    "cjklist13.rtf",
    "cjklist16.rtf",
    "cjklist20.rtf",
    "cjklist21.rtf",
    "column-break.rtf",
    "hidden-linebreaks.rtf",
    "hyperlink-target.rtf",
    "hyperlink-with-backslashes.rtf",
    "hyperlink.rtf",
    "hyperlink_empty.rtf",
    "page-background.rtf",
    "page-border.rtf",
    "page-break-emptyparas-spltpgpar.rtf",
    "page-break-emptyparas.rtf",
    "para-adjust-distribute.rtf",
    "para-border.rtf",
    "para-bottom-margin.rtf",
    "para-shadow.rtf",
    "para-style-bottom-margin-2.rtf",
    "test437Encoding.rtf",
    "test874Encoding.rtf",
    "test950Encoding.rtf",
    "testDefaultEncodingParse.rtf",
    "testEncodingParse.rtf",
    "testGreekEncoding.rtf",
    "testHex.rtf",
    "testJapaneseJisEncoding.rtf",
    "testJapaneseJisEncodingTwoFonts.rtf",
    "testJapaneseUtf8Encoding.rtf",
    "testKoreanEncoding.rtf",
    "testMultiByteHex.rtf",
    "testNecCharacters.rtf",
    "testNegativeUnicode.rtf",
    "testSpecialChars.rtf",
    "testStyles.rtf",
    "testTurkishEncoding.rtf",
    "testUnicode.rtf",
    "testUpr.rtf",
    "watermark.rtf",
];

const UNSUPPORTED_EXACT_CODECS: &[&str] = &["test10001Encoding.rtf", "test10007Encoding.rtf"];

#[test]
fn parses_the_real_world_rtf_compatibility_corpus() {
    let corpus = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/rtf");
    for fixture in FIXTURES {
        let path = corpus.join(fixture);
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        Tree::from_bytes(&bytes, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to parse {}: {error}", path.display()));
    }

    // Mac Japanese 10001 and Mac Russian 10007 need distinct codecs. Treating
    // them as Windows Shift-JIS or KOI8-R would be silent data corruption.
    for fixture in UNSUPPORTED_EXACT_CODECS {
        let path = corpus.join(fixture);
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        // Refused typed as an unsupported code page, not as malformed
        // (record 0347).
        assert!(
            matches!(
                Tree::from_bytes(&bytes, litchi_rtf::read::OpenOptions::default()),
                Err(litchi_rtf::Error::UnsupportedCodePage(10001 | 10007))
            ),
            "{} is not refused as an unsupported code page",
            path.display()
        );
    }
}
