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
use litchi_rtf::{DocumentPrivacyPolicies, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_both_privacy_requests_without_removing_metadata_or_body() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1\rempersonalinfo\remdttm"#,
            r#"{\info{\author Alice}{\creatim\yr2020\mo1\dy2}}Body}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        *document.privacy_policies(),
        DocumentPrivacyPolicies {
            remove_personal_information: true,
            remove_date_time_information: true,
        }
    );
    assert_eq!(document.text(), "Body");
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(serialized.contains("Alice"));
    assert!(serialized.contains("\\creatim"));
}

#[test]
fn parses_each_request_independently() {
    for (name, personal, date_time) in [("rempersonalinfo", true, false), ("remdttm", false, true)]
    {
        let source = format!(r"{{\rtf1\{name} Body}}");
        let document = Tree::parse(&source, litchi_rtf::read::OpenOptions::default()).unwrap();
        assert_eq!(
            document.privacy_policies().remove_personal_information,
            personal
        );
        assert_eq!(
            document.privacy_policies().remove_date_time_information,
            date_time
        );
        assert_eq!(document.text(), "Body");
    }
}

#[test]
fn omission_is_false_and_serializes_nothing() {
    let document = Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(document.privacy_policies().is_empty());
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("rempersonalinfo"));
    assert!(!serialized.contains("remdttm"));
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_passively() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    document.set_privacy_policies(DocumentPrivacyPolicies {
        remove_personal_information: true,
        remove_date_time_information: true,
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\rempersonalinfo").unwrap() < serialized.find("\\remdttm").unwrap());
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.privacy_policies(), document.privacy_policies());
    assert_eq!(reparsed.text(), "Body");

    document.clear_privacy_policies();
    assert!(document.privacy_policies().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_bundled_libreoffice_privacy_policy_producer_fixture() {
    let fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/core/data/rtf/pass/tdf116851.rtf"
    );
    let document = Tree::parse_bytes(fixture, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(document.privacy_policies().remove_personal_information);
    assert!(document.privacy_policies().remove_date_time_information);
    assert!(!document.text().is_empty());
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_requests() {
    for name in ["rempersonalinfo", "remdttm"] {
        for suffix in ["0", "1", "-1", "2147483647", "99999999999"] {
            let source = format!(r"{{\rtf1\{name}{suffix} Body}}");
            assert!(
                repair_support::rejected_or_parameter_repaired(&Tree::parse(
                    &source,
                    litchi_rtf::read::OpenOptions::default()
                )),
                "accepted malformed {source}"
            );
        }
        for source in [
            format!(r"{{\rtf1\{name}\{name} Body}}"),
            format!(r"{{\rtf1{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1{{\{name}}}Body}}"),
            format!(r"{{\rtf1 Body\{name}}}"),
        ] {
            assert!(
                repair_support::rejected_or_parameter_repaired(&Tree::parse(
                    &source,
                    litchi_rtf::read::OpenOptions::default()
                )),
                "accepted malformed {source}"
            );
        }
    }
}
