use super::{Alternatives, Branch, Limits, read};
use crate::mce::Capabilities;

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

fn source() -> &'static [u8] {
    br#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:word="urn:word" xmlns:vendor="urn:vendor">
  <mc:Choice Requires="vendor"><vendor:future a="1"><!--kept--><vendor:nested/></vendor:future></mc:Choice>
  <mc:Choice Requires="word"><word:known/></mc:Choice>
  <mc:Fallback><word:fallback/></mc:Fallback>
</mc:AlternateContent>"#
}

fn snapshot() -> Alternatives {
    read(source(), &Limits::default()).expect("valid AlternateContent")
}

#[test]
fn retains_one_exact_source_and_exposes_typed_branches() {
    let alternatives = snapshot();
    assert_eq!(alternatives.as_xml(), source());
    assert_eq!(alternatives.len(), 3);

    let branches = alternatives.branches().collect::<Vec<_>>();
    assert_eq!(branches.len(), 3);
    let Branch::Choice(choice) = &branches[0] else {
        panic!("first branch must be a Choice");
    };
    assert_eq!(
        choice.requirements().collect::<Vec<_>>(),
        vec!["urn:vendor"]
    );
    assert_eq!(
        choice.content(),
        br#"<vendor:future a="1"><!--kept--><vendor:nested/></vendor:future>"#
    );
    assert_eq!(branches[2].content(), br#"<word:fallback/>"#);
}

#[test]
fn selects_first_supported_choice_then_fallback_without_discarding_inactive_data() {
    let alternatives = snapshot();
    let mut capabilities = Capabilities::new();
    capabilities.understand_namespace("urn:word");

    let Branch::Choice(choice) = alternatives
        .select(&capabilities)
        .expect("word choice is supported")
    else {
        panic!("expected supported Choice");
    };
    assert_eq!(choice.content(), br#"<word:known/>"#);
    assert!(
        std::str::from_utf8(alternatives.as_xml())
            .expect("source is UTF-8")
            .contains("vendor:future")
    );

    let alternatives = snapshot();
    let fallback = alternatives.select(&Capabilities::new()).expect("fallback");
    assert!(matches!(fallback, Branch::Fallback(_)));
    assert_eq!(fallback.content(), br#"<word:fallback/>"#);
}

#[test]
fn accepts_nested_markup_and_arbitrary_prefixes() {
    let xml = format!(
        r#"<x:AlternateContent xmlns:x="{MC}" xmlns:f="urn:feature"><x:Choice Requires="f"><x:AlternateContent><x:Choice Requires="f"><f:value/></x:Choice></x:AlternateContent></x:Choice></x:AlternateContent>"#
    );
    let alternatives = read(xml.as_bytes(), &Limits::default()).expect("nested MCE");
    assert_eq!(alternatives.as_xml(), xml.as_bytes());
    assert_eq!(alternatives.len(), 1);

    let local = format!(
        r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice xmlns:f="urn:feature" Requires="f"><f:value/></mc:Choice></mc:AlternateContent>"#
    );
    let local = read(local.as_bytes(), &Limits::default()).expect("local branch namespace");
    assert_eq!(
        local.branch(0).expect("local choice").as_xml(),
        br#"<mc:Choice xmlns:f="urn:feature" Requires="f"><f:value/></mc:Choice>"#
    );
}

#[test]
fn rejects_invalid_branch_grammar_and_unsafe_markup() {
    let invalid = [
        format!(r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice/></mc:AlternateContent>"#),
        format!(
            r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Fallback/><mc:Choice Requires="mc"/></mc:AlternateContent>"#
        ),
        format!(
            r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Fallback/><mc:Fallback/></mc:AlternateContent>"#
        ),
        format!(
            r#"<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="missing"><x:value/></mc:Choice></mc:AlternateContent>"#
        ),
        format!(
            r#"<!DOCTYPE root><mc:AlternateContent xmlns:mc="{MC}"><mc:Fallback/></mc:AlternateContent>"#
        ),
        format!(r#"<mc:AlternateContent xmlns:mc="{MC}"/>"#),
    ];
    for xml in invalid {
        assert!(read(xml.as_bytes(), &Limits::default()).is_err(), "{xml}");
    }
}

#[test]
fn enforces_byte_depth_node_and_branch_budgets() {
    let input = source();
    assert!(
        read(
            input,
            &Limits {
                bytes: input.len() - 1,
                ..Limits::default()
            }
        )
        .is_err()
    );
    assert!(
        read(
            input,
            &Limits {
                depth: 1,
                ..Limits::default()
            }
        )
        .is_err()
    );
    assert!(
        read(
            input,
            &Limits {
                nodes: 1,
                ..Limits::default()
            }
        )
        .is_err()
    );
    assert!(
        read(
            input,
            &Limits {
                branches: 2,
                ..Limits::default()
            }
        )
        .is_err()
    );
}
