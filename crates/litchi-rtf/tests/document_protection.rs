#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::write::Writer;
use litchi_rtf::{
    Document, MAX_PASSWORD_HASH_BYTES, PasswordHash, ProtectionLevel, ProtectionType, RtfDocument,
    RtfWriter,
};
use std::borrow::Cow;
use std::fs;

const MODERN_PASSWORD_HASH_SAMPLE: &str = "010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49";

fn decode_hex(value: &str) -> Vec<u8> {
    value
        .as_bytes()
        .chunks_exact(2)
        .map(|pair| {
            let high = (pair[0] as char).to_digit(16).unwrap() as u8;
            let low = (pair[1] as char).to_digit(16).unwrap() as u8;
            (high << 4) | low
        })
        .collect()
}

fn round_trip(document: &RtfDocument<'_>) -> RtfDocument<'static> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    RtfDocument::parse_bytes(&output).unwrap()
}

#[test]
fn parses_all_protection_controls_and_round_trips_inert_hash() {
    let source = concat!(
        r#"{\rtf1\ansi{\info{\title Protected}{\*\password aBcD0123}}"#,
        r#"\formprot\annotprot0\revprot\readprot0\allprot"#,
        r#"\enforceprot1\protlevel2 Body}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let protection = document.protection();
    assert_eq!(protection.forms, Some(true));
    assert_eq!(protection.annotations, Some(false));
    assert_eq!(protection.revisions, Some(true));
    assert_eq!(protection.read_only, Some(false));
    assert_eq!(protection.all, Some(true));
    assert_eq!(protection.enforced, Some(true));
    assert_eq!(protection.level, Some(ProtectionLevel::Level2));
    assert_eq!(protection.password_hash.as_deref(), Some("aBcD0123"));
    assert_eq!(
        protection.protection_type(),
        ProtectionType::RevisionTracking
    );
    assert_eq!(document.text(), "Body");

    let reparsed = round_trip(&document);
    assert_eq!(reparsed.protection(), protection);
    assert_eq!(reparsed.info().title, document.info().title);
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_malformed_duplicate_or_misplaced_protection() {
    for source in [
        r"{\rtf1\formprot2}",
        r"{\rtf1\enforceprot}",
        r"{\rtf1\enforceprot2}",
        r"{\rtf1\protlevel}",
        r"{\rtf1\protlevel4}",
        r"{\rtf1\formprot\formprot}",
        r"{\rtf1{\b\readprot}}",
        r"{\rtf1 Body\revprot}",
        r"{\rtf1{\info{\password 00000000}}}",
        r"{\rtf1{\info{\*\password xyz}}}",
        r"{\rtf1{\info{\*\password 00000000}{\*\password 11111111}}}",
        r"{\rtf1{\*\passwordhash1 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}}",
        r"{\rtf1{\info{\*\passwordhash 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}}}",
        r"{\rtf1\sect{\*\passwordhash 010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49}}",
        &format!(
            r"{{\rtf1{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}}{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}}}}"
        ),
        &format!(r"{{\rtf1{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}Z}}}}"),
        &format!(r"{{\rtf1{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}0}}}}"),
        &format!(r"{{\rtf1{{\*\passwordhash {{{MODERN_PASSWORD_HASH_SAMPLE}}}}}}}"),
        &format!(r"{{\rtf1{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}\b}}}}"),
        &format!(
            r"{{\rtf1\trowd\cellx1000\intbl\cell\row{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}}}}"
        ),
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn parses_real_libreoffice_protection_fixtures() {
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras"
    );

    let read_only = fs::read(format!("{root}/rtfimport/data/read-only-protect.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&read_only).unwrap();
    assert_eq!(document.protection().annotations, Some(true));
    assert_eq!(document.protection().read_only, Some(true));
    assert_eq!(document.protection().enforced, Some(true));
    assert_eq!(document.protection().level, Some(ProtectionLevel::Level3));

    let forms = fs::read(format!("{root}/rtfexport/data/4010_min.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&forms).unwrap();
    assert_eq!(document.protection().forms, Some(true));
    assert_eq!(document.protection().all, Some(true));
    assert_eq!(document.protection().level, Some(ProtectionLevel::Level2));

    let password = fs::read(format!("{root}/rtfexport/data/fdo55504-1-min.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&password).unwrap();
    assert_eq!(
        document.protection().password_hash.as_deref(),
        Some("00000000")
    );
}

#[test]
fn parses_modern_passwordhash_as_inert_typed_payload() {
    // RTF 1.9.1, Document Properties, p. 40 example vector.  The record is
    // checked as data only; this test never derives or verifies a password.
    let source = format!(
        r#"{{\rtf1\ansi{{\*\passwordhash
{MODERN_PASSWORD_HASH_SAMPLE}}} Body}}"#
    );
    let document = RtfDocument::parse(&source).unwrap();
    let hash = document.protection().modern_password_hash().unwrap();
    assert_eq!(hash.version(), 1);
    assert_eq!(hash.total_size(), 76);
    assert_eq!(hash.algorithm_id(), 0x8004);
    assert_eq!(hash.spin_count(), 50_000);
    assert_eq!(hash.salt_size(), 16);
    assert_eq!(hash.hash_size(), 20);
    assert_eq!(hash.bytes().len(), 76);
    assert_eq!(
        hash.bytes().get(..28),
        Some(
            &[
                1, 0, 0, 0, 76, 0, 0, 0, 1, 0, 0, 0, 4, 128, 0, 0, 80, 195, 0, 0, 20, 0, 0, 0, 16,
                0, 0, 0,
            ][..]
        )
    );
    assert_eq!(
        hash.salt(),
        &[
            248, 156, 54, 13, 12, 157, 54, 13, 0, 0, 0, 0, 139, 194, 158, 47
        ]
    );
    assert_eq!(
        hash.hash(),
        &[
            120, 162, 20, 65, 34, 237, 104, 161, 112, 30, 46, 165, 11, 187, 190, 175, 115, 51, 196,
            13
        ]
    );
    assert_eq!(
        hash.bytes().get(64..),
        Some(&[254, 4, 140, 207, 85, 247, 9, 184, 204, 126, 139, 73][..])
    );
    let reparsed = round_trip(&document);
    assert_eq!(reparsed.protection(), document.protection());
    assert!(reparsed.protection().password_hash.is_none());
}

#[test]
fn preserves_unknown_passwordhash_bytes_and_enforces_bounds() {
    let mut bytes = vec![
        1, 0, 0, 0, 32, 0, 0, 0, 0x78, 0x56, 0x34, 0x12, 4, 3, 2, 1, 1, 0, 0, 0, 1, 0, 0, 0, 1, 0,
        0, 0, 0xAA, 0xBB, 0xCC, 0xDD,
    ];
    let value = PasswordHash::from_bytes(Cow::Borrowed(&bytes)).unwrap();
    assert_eq!(value.flags(), 0x12345678);
    assert_eq!(value.to_bytes(), bytes);
    bytes.extend(std::iter::repeat_n(0, MAX_PASSWORD_HASH_BYTES));
    assert!(PasswordHash::from_bytes(Cow::Owned(bytes)).is_err());
}

#[test]
fn immutable_facade_write_preserves_passwordhash_source_exactly() {
    let source = concat!(
        r#"{\rtf1\ansi {\*\passwordhash"#,
        "\n010000004c000000010000000480000050c300001400000010000000f89c360d0c9d360d000000008bc29e2f78a2144122ed68a1701e2ea50bbbbeaf7333c40dfe048ccf55f709b8cc7e8b49",
        "}\nBody}"
    );
    let document = Document::from_bytes(source.as_bytes()).unwrap();
    let mut output = Vec::new();
    Writer::new(&mut output).write(&document).unwrap();
    assert_eq!(output, source.as_bytes());
    assert_eq!(Document::from_bytes(&output).unwrap().text(), "Body");
}

#[test]
fn authored_passwordhash_removal_and_readdition_round_trip_semantically() {
    let source = format!(
        r#"{{\rtf1\ansi{{\*\passwordhash
{MODERN_PASSWORD_HASH_SAMPLE}}} Body}}"#
    );
    let mut document = RtfDocument::parse(&source).unwrap();
    let original = document.protection().clone().into_owned();
    document.clear_protection();
    assert!(
        round_trip(&document)
            .protection()
            .modern_password_hash()
            .is_none()
    );
    document.set_protection(original.clone()).unwrap();
    assert_eq!(round_trip(&document).protection(), &original);
}

#[test]
fn parser_rejects_passwordhash_hex_beyond_the_decoded_bound() {
    let payload = "00".repeat(MAX_PASSWORD_HASH_BYTES + 1);
    let source = format!(r"{{\rtf1{{\*\passwordhash {payload}}}}}");
    assert!(RtfDocument::parse(&source).is_err());
}

#[test]
fn rejects_passwordhash_after_structural_body_events() {
    let field = format!(
        r"{{\rtf1\ansi{{\field{{\*\fldinst X}}{{\fldrslt }}}}{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}} Body}}"
    );
    assert!(RtfDocument::parse(&field).is_err());

    for prefix in [r"\par", r"\line", r"\page"] {
        let source = format!(
            r"{{\rtf1\ansi{prefix}{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}} Body}}"
        );
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }
}

#[test]
fn rejects_passwordhash_after_zero_width_body_structures() {
    let prefixes = [
        r#"{\rtf1\ansi{\*\bkmkstart bm}{\*\bkmkend bm}"#,
        r#"{\rtf1\ansi{\footnote N}"#,
        r#"{\rtf1\ansi{\object\objemb{\*\objdata 00}}"#,
        r#"{\rtf1\ansi{\pict\pngblip 00}"#,
        r#"{\rtf1\ansi{\header H}"#,
    ];
    for prefix in prefixes {
        let valid = format!(r"{prefix} Body}}");
        assert!(
            RtfDocument::parse(&valid).is_ok(),
            "invalid control fixture {valid}"
        );
        let source = format!(r"{prefix}{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}}} Body}}");
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }
}

#[test]
fn rejects_nested_known_passwordhash_before_opaque_or_specialized_parsing() {
    let unknown =
        format!(r"{{\rtf1\ansi{{\*\unknown{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}Z}}}}}}");
    let field = format!(
        r"{{\rtf1\ansi{{\field{{\*\fldinst X}}{{\fldrslt {{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}Z}}}}}}}}"
    );
    let header =
        format!(r"{{\rtf1\ansi{{\header{{\*\passwordhash {MODERN_PASSWORD_HASH_SAMPLE}Z}}}}}}");
    for source in [unknown, field, header] {
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }
}

#[test]
fn rejects_impossible_declared_passwordhash_size_at_header_boundary() {
    let source = r"{\rtf1\ansi{\*\passwordhash 01000000FFFFFFFF}}";
    assert!(RtfDocument::parse(source).is_err());
}

#[test]
fn rejects_passwordhash_bytes_after_declared_total_before_push() {
    let mut header = Vec::new();
    for value in [1_u32, 28, 0, 0, 0, 0, 0] {
        header.extend_from_slice(&value.to_le_bytes());
    }
    let header_hex = header
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect::<String>();
    let complete_extra = format!(r"{{\rtf1\ansi{{\*\passwordhash {header_hex}00}}}}");
    let odd_extra = format!(r"{{\rtf1\ansi{{\*\passwordhash {header_hex}0}}}}");

    assert!(RtfDocument::parse(&complete_extra).is_err());
    assert!(RtfDocument::parse(&odd_extra).is_err());
}

#[test]
fn authored_layout_can_reproduce_the_full_rtf_sample_tail() {
    let sample = decode_hex(MODERN_PASSWORD_HASH_SAMPLE);
    let value = PasswordHash::from_parts(
        1,
        1,
        0x8004,
        50_000,
        Cow::Owned(sample[28..44].to_vec()),
        Cow::Owned(sample[44..64].to_vec()),
        Cow::Owned(sample[64..].to_vec()),
    )
    .unwrap();
    assert_eq!(value.bytes(), sample.as_slice());
}
