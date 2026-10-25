#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

mod repair_support;

use litchi_rtf::raw::Tree;
use litchi_rtf::{DocumentPrintLayoutSettings, MAX_DOCUMENT_GUTTER_TWIPS, write::Writer};

fn write(document: &Tree<'_>) -> Result<Vec<u8>, std::io::Error> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document)?;
    Ok(output)
}

#[test]
fn parses_toggle_flag_and_global_gutter() {
    let document = Tree::parse(
        r"{\rtf1\facingp\margmirror\gutter720\gutterprl\twoonone Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        *document.print_layout_settings(),
        DocumentPrintLayoutSettings {
            facing_pages: true,
            mirror_margins: true,
            document_gutter_twips: Some(720),
            parallel_gutter: true,
            two_logical_pages_per_physical_page: true,
        }
    );

    let disabled = Tree::parse(
        r"{\rtf1\facingp0 Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(!disabled.print_layout_settings().facing_pages);
}

#[test]
fn global_gutter_is_inherited_and_guttersxn_overrides_after_reset() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1\gutter720\sectd\guttersxn360 First"#,
            r#"\sect\sectd Second\sect\sectd\guttersxn0 Third}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.sections().len(), 3);
    assert_eq!(document.sections()[0].properties.margin_gutter, 360);
    assert_eq!(document.sections()[1].properties.margin_gutter, 720);
    assert_eq!(document.sections()[2].properties.margin_gutter, 0);

    let serializable = Tree::parse(
        r"{\rtf1\gutter720\sectd\guttersxn360 First}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let reparsed = Tree::parse_bytes(
        &write(&serializable).unwrap(),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        reparsed.print_layout_settings(),
        serializable.print_layout_settings()
    );
    assert_eq!(reparsed.sections()[0].properties.margin_gutter, 360);
}

#[test]
fn public_setters_are_atomic_and_writer_rejects_direct_invalid_values() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    let valid = DocumentPrintLayoutSettings {
        facing_pages: true,
        mirror_margins: true,
        document_gutter_twips: Some(MAX_DOCUMENT_GUTTER_TWIPS),
        parallel_gutter: true,
        two_logical_pages_per_physical_page: true,
    };
    document.set_print_layout_settings(valid).unwrap();
    assert!(
        document
            .set_document_gutter_twips(Some(MAX_DOCUMENT_GUTTER_TWIPS + 1))
            .is_err()
    );
    assert_eq!(*document.print_layout_settings(), valid);

    let mut settings = valid;
    assert!(
        settings
            .set_document_gutter_twips(Some(MAX_DOCUMENT_GUTTER_TWIPS + 1))
            .is_err()
    );
    assert_eq!(settings, valid);

    let invalid = DocumentPrintLayoutSettings {
        document_gutter_twips: Some(MAX_DOCUMENT_GUTTER_TWIPS + 1),
        ..valid
    };
    assert!(document.set_print_layout_settings(invalid).is_err());
    assert_eq!(*document.print_layout_settings(), valid);

    document.print_layout_settings().validate().unwrap();
    let mut invalid_document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut invalid_settings = *invalid_document.print_layout_settings();
    invalid_settings.document_gutter_twips = Some(MAX_DOCUMENT_GUTTER_TWIPS + 1);
    assert!(
        invalid_document
            .set_print_layout_settings(invalid_settings)
            .is_err()
    );

    let mut output = Vec::new();
    assert!(
        Writer::new(&mut output)
            .write_document_print_layout_settings(&invalid)
            .is_err()
    );
}

#[test]
fn writer_uses_canonical_order_and_round_trips() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    document
        .set_print_layout_settings(DocumentPrintLayoutSettings {
            facing_pages: true,
            mirror_margins: true,
            document_gutter_twips: Some(720),
            parallel_gutter: true,
            two_logical_pages_per_physical_page: true,
        })
        .unwrap();
    let output = write(&document).unwrap();
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\facingp",
        "\\margmirror",
        "\\gutter720",
        "\\gutterprl",
        "\\twoonone",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.print_layout_settings(),
        document.print_layout_settings()
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_malformed_bounds_duplicates_grouping_and_late_placement() {
    for source in [
        r"{\rtf1\gutter Body}",
        r"{\rtf1\gutter-1 Body}",
        r"{\rtf1\gutter31681 Body}",
        r"{\rtf1\gutter99999999999 Body}",
        r"{\rtf1\margmirror0 Body}",
        r"{\rtf1\facingp\facingp0 Body}",
        r"{\rtf1\margmirror\margmirror Body}",
        r"{\rtf1\gutter1\gutter2 Body}",
        r"{\rtf1{\facingp}Body}",
        r"{\rtf1{\margmirror}Body}",
        r"{\rtf1{\gutter720}Body}",
        r"{\rtf1 Body\facingp}",
        r"{\rtf1 Body\margmirror}",
        r"{\rtf1 Body\gutter720}",
    ] {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                source,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn destination_controls_never_leak_into_document_settings() {
    let ignored = Tree::parse(
        r"{\rtf1{\*\facingp}Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(ignored.print_layout_settings().is_empty());
    assert_eq!(ignored.text(), "Body");

    for source in [
        r"{\rtf1{\header\facingp\margmirror\gutter720 Header}Body}",
        r"{\rtf1{\footer\facingp\margmirror\gutter720 Footer}Body}",
        r"{\rtf1{\footnote\facingp\margmirror\gutter720 Note}Body}",
        r"{\rtf1{\*\unknown\facingp\margmirror\gutter720 Hidden}Body}",
    ] {
        if let Ok(document) = Tree::parse(source, litchi_rtf::read::OpenOptions::default()) {
            assert!(
                document.print_layout_settings().is_empty(),
                "leaked {source}"
            );
        }
    }
}

#[test]
fn bundled_libreoffice_fixtures_round_trip_layout_settings() {
    const FIXTURES: &[(&str, &[u8])] = &[
        (
            "margmirror.rtf",
            include_bytes!(
                "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/margmirror.rtf"
            ),
        ),
        (
            "gutter-left.rtf",
            include_bytes!(
                "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/gutter-left.rtf"
            ),
        ),
        (
            "rhbz1065629.rtf",
            include_bytes!(
                "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/rhbz1065629.rtf"
            ),
        ),
    ];

    for (name, source) in FIXTURES {
        let document = Tree::parse_bytes(source, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to parse {name}: {error}"));
        let settings = *document.print_layout_settings();
        match *name {
            "margmirror.rtf" => assert!(settings.mirror_margins),
            "gutter-left.rtf" => assert_eq!(settings.document_gutter_twips, Some(720)),
            "rhbz1065629.rtf" => assert!(settings.facing_pages),
            _ => unreachable!(),
        }
        let output =
            write(&document).unwrap_or_else(|error| panic!("failed to write {name}: {error}"));
        let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to reparse {name}: {error}"));
        assert_eq!(*reparsed.print_layout_settings(), settings, "{name}");
    }
}
