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
use litchi_rtf::write::{Options, Writer};
use litchi_rtf::{DEFAULT_TAB_WIDTH_TWIPS, DefaultTabWidthPolicy, MAX_DEFAULT_TAB_WIDTH_TWIPS};

fn write(doc: &Tree<'_>, default_tab_width: DefaultTabWidthPolicy) -> std::io::Result<String> {
    let mut bytes = Vec::new();
    let options = Options {
        default_tab_width,
        ..Options::default()
    };
    Writer::with_options(&mut bytes, options).write_document(doc)?;
    Ok(String::from_utf8(bytes).expect("writer emits ASCII-compatible RTF"))
}

#[test]
fn omission_and_explicit_producer_values_remain_distinct() {
    let omitted = Tree::parse(
        r"{\rtf1\ansi body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(omitted.default_tab_width_twips(), None);
    assert_eq!(
        omitted.effective_default_tab_width_twips(),
        DEFAULT_TAB_WIDTH_TWIPS
    );

    for width in [0, 14, 284, 420, 480, 708, 709, 720, 840, 1298, i32::MAX] {
        let doc = Tree::parse(
            &format!(r"{{\rtf1\ansi\deftab{width} body}}"),
            litchi_rtf::read::OpenOptions::default(),
        )
        .unwrap();
        assert_eq!(doc.default_tab_width_twips(), Some(width.cast_unsigned()));
        assert_eq!(
            doc.effective_default_tab_width_twips(),
            width.cast_unsigned()
        );
    }
}

#[test]
fn preserve_policy_roundtrips_presence_and_value() {
    let omitted = Tree::parse(
        r"{\rtf1\ansi body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let omitted_rtf = write(&omitted, DefaultTabWidthPolicy::PreserveDocument).unwrap();
    assert!(!omitted_rtf.contains("\\deftab"));
    assert_eq!(
        Tree::parse(&omitted_rtf, litchi_rtf::read::OpenOptions::default())
            .unwrap()
            .default_tab_width_twips(),
        None
    );

    let explicit = Tree::parse(
        r"{\rtf1\ansi\deftab708 body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let explicit_rtf = write(&explicit, DefaultTabWidthPolicy::PreserveDocument).unwrap();
    assert!(explicit_rtf.contains("\\deftab708"));
    assert_eq!(
        Tree::parse(&explicit_rtf, litchi_rtf::read::OpenOptions::default())
            .unwrap()
            .default_tab_width_twips(),
        Some(708)
    );
}

#[test]
fn override_policy_is_the_only_writer_precedence_over_document_metadata() {
    let source = Tree::parse(
        r"{\rtf1\ansi\deftab708 body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let overridden = write(&source, DefaultTabWidthPolicy::Override(480)).unwrap();
    assert!(overridden.contains("\\deftab480"));
    assert!(!overridden.contains("\\deftab708"));

    let omitted = Tree::parse(
        r"{\rtf1\ansi body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(
        write(&omitted, DefaultTabWidthPolicy::Override(720))
            .unwrap()
            .contains("\\deftab720")
    );
    assert!(
        write(&omitted, DefaultTabWidthPolicy::Override(0))
            .unwrap()
            .contains("\\deftab0")
    );
}

#[test]
fn typed_mutation_preserves_passivity_and_can_restore_omission() {
    let mut doc = Tree::parse(
        r"{\rtf1\ansi body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    doc.set_default_tab_width_twips(840).unwrap();
    let serialized = write(&doc, DefaultTabWidthPolicy::PreserveDocument).unwrap();
    assert!(serialized.contains("\\deftab840"));
    assert!(!serialized.contains("\\tx"));
    assert!(!serialized.contains("\\tb"));

    doc.clear_default_tab_width();
    assert_eq!(doc.default_tab_width_twips(), None);
    assert!(
        !write(&doc, DefaultTabWidthPolicy::PreserveDocument)
            .unwrap()
            .contains("\\deftab")
    );
    assert!(
        doc.set_default_tab_width_twips(MAX_DEFAULT_TAB_WIDTH_TWIPS + 1)
            .is_err()
    );
}

#[test]
fn malformed_domain_duplicates_and_placement_are_rejected() {
    for hostile in [
        r"{\rtf1\ansi\deftab body}",
        r"{\rtf1\ansi\deftab-1 body}",
        r"{\rtf1\ansi\deftab2147483648 body}",
        r"{\rtf1\ansi\deftab720\deftab480 body}",
        r"{\rtf1\ansi{\deftab720}body}",
        r"{\rtf1\ansi{\*\deftab720}body}",
        r"{\rtf1\ansi body\deftab720}",
    ] {
        assert!(
            repair_support::rejected_or_parameter_repaired(&Tree::parse(
                hostile,
                litchi_rtf::read::OpenOptions::default()
            )),
            "accepted {hostile}"
        );
    }

    let doc = Tree::parse(
        r"{\rtf1\ansi body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert!(
        write(
            &doc,
            DefaultTabWidthPolicy::Override(MAX_DEFAULT_TAB_WIDTH_TWIPS + 1)
        )
        .is_err()
    );
}

#[test]
fn libreoffice_fixture_exposes_its_explicit_default_tab_width() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/core/data/rtf/pass/fdo78900.rtf"
    );
    let bytes = std::fs::read(path).unwrap();
    let doc = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(doc.default_tab_width_twips(), Some(720));
}

#[test]
fn identical_repeated_declarations_are_idempotent() {
    // LibreOffice restates `\deftab` once per paragraph-properties reset, so an
    // identical redeclaration carries no new information and must not discard
    // the document.
    let doc = Tree::parse(
        r"{\rtf1\ansi\deftab720\deftab720\deftab720 body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .expect("repeated identical deftab rejected");
    assert_eq!(doc.default_tab_width_twips(), Some(720));
    assert_eq!(doc.text().trim(), "body");
}

#[test]
fn conflicting_repeated_declarations_are_still_rejected() {
    let Err(error) = Tree::parse(
        r"{\rtf1\ansi\deftab720\deftab480 body}",
        litchi_rtf::read::OpenOptions::default(),
    ) else {
        panic!("conflicting deftab accepted");
    };
    assert!(
        error.to_string().contains("conflicting RTF deftab"),
        "unexpected error: {error}"
    );
}

#[test]
fn libreoffice_fixtures_with_repeated_deftab_keep_their_text() {
    // Both fixtures are genuine LibreOffice output that repeats `\deftab`; each
    // was previously rejected outright, losing all of its body text.
    for (fixture, expected_twips, expected_text) in [
        ("repeated-deftab.rtf", 420u32, "standard"),
        ("repeated-deftab-cjk.rtf", 840u32, "のをする。"),
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/rtf")
            .join(fixture);
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        let doc = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default())
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert_eq!(doc.default_tab_width_twips(), Some(expected_twips));
        assert!(
            doc.text().contains(expected_text),
            "{fixture} lost its body text"
        );
    }
}
