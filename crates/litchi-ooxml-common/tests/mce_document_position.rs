//! Document-position rules of the two Markup Compatibility processors: what
//! may appear before, inside and after the root element, and how each
//! processor reports a violation.

use litchi_ooxml_common::mce::{
    Capabilities, Error, Limits, SemanticEvent, StreamError, StreamLimits,
    process_markup_compatibility, process_markup_compatibility_stream_with_observers,
};
use std::io::Cursor;

const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// A root whose content fails semantic processing with a recoverable error,
/// which moves the stream processor to its raw-only flow.
fn failing_root_open() -> String {
    format!(
        r#"<r xmlns:mc="{MC}"><mc:AlternateContent><mc:Choice Requires="missing"/></mc:AlternateContent>"#
    )
}

/// A complete root element that fails `MustUnderstand` on its own start tag.
fn failing_empty_root() -> String {
    format!(r#"<r xmlns:mc="{MC}" xmlns:x="urn:x" mc:MustUnderstand="x"/>"#)
}

fn describe(error: &Error) -> String {
    match error {
        Error::NonConformant(message) => format!("NonConformant({message})"),
        Error::MustUnderstand(message) => format!("MustUnderstand({message})"),
        Error::Xml(_) => "Xml".to_owned(),
        other => format!("{other:?}"),
    }
}

/// The outcome of the stream processor, with the semantic events it delivered
/// counted by kind: declarations, element starts and ends, everything else.
fn stream(xml: &str) -> String {
    let mut declarations = 0usize;
    let mut elements = 0usize;
    let mut texts = 0usize;
    let mut raw_elements = 0usize;
    let mut input = Cursor::new(xml.as_bytes());
    let result = process_markup_compatibility_stream_with_observers(
        &mut input,
        &Capabilities::new(),
        &StreamLimits::default(),
        |_| {
            raw_elements += 1;
            Ok::<(), std::convert::Infallible>(())
        },
        |event| {
            match event {
                SemanticEvent::Decl(_) => declarations += 1,
                SemanticEvent::Start(_) | SemanticEvent::Empty(_) | SemanticEvent::End(_) => {
                    elements += 1;
                },
                _ => texts += 1,
            }
            Ok::<(), std::convert::Infallible>(())
        },
    );
    let outcome = match result {
        Ok(_) => "ok".to_owned(),
        Err(StreamError::Mce {
            error,
            prior_mce_error,
            ..
        }) => match prior_mce_error {
            Some(prior) => format!("{} after {}", describe(&error), describe(&prior)),
            None => describe(&error),
        },
        Err(StreamError::Input { .. }) => "input".to_owned(),
        Err(_) => "callback".to_owned(),
    };
    format!("{outcome} [decl={declarations} elem={elements} text={texts} raw={raw_elements}]")
}

fn buffered(xml: &str) -> Result<String, String> {
    process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &Limits::default())
        .map(|output| String::from_utf8(output.xml.into_owned()).unwrap())
        .map_err(|error| describe(&error))
}

fn assert_outcomes(cases: &[(String, &str)]) {
    let mismatches: Vec<String> = cases
        .iter()
        .filter_map(|(xml, expected)| {
            let found = stream(xml);
            (found != *expected)
                .then(|| format!("{xml:?}\n   found: {found}\nexpected: {expected}"))
        })
        .collect();
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

#[test]
fn stream_accepts_a_prolog_one_root_and_an_epilog() {
    assert_outcomes(&[
        ("<r/>".to_owned(), "ok [decl=0 elem=1 text=0 raw=1]"),
        ("<r></r>".to_owned(), "ok [decl=0 elem=2 text=0 raw=1]"),
        (r#"<?xml version="1.0"?><r/>"#.to_owned(), "ok [decl=1 elem=1 text=0 raw=1]"),
        (
            "<?xml version=\"1.0\"?>\n<!--before-->\n<r>t<a/>&amp;<![CDATA[c]]></r>\n<!--after-->\n"
                .to_owned(),
            "ok [decl=1 elem=3 text=3 raw=2]",
        ),
        ("<!--before--><r/>".to_owned(), "ok [decl=0 elem=1 text=0 raw=1]"),
        ("\n<r/>\n".to_owned(), "ok [decl=0 elem=1 text=0 raw=1]"),
    ]);
}

#[test]
fn stream_requires_exactly_one_root() {
    assert_outcomes(&[
        (
            String::new(),
            "NonConformant(missing document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            "  \n".to_owned(),
            "NonConformant(missing document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            "<!--only-->".to_owned(),
            "NonConformant(missing document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            r#"<?xml version="1.0"?>"#.to_owned(),
            "NonConformant(missing document root) [decl=1 elem=0 text=0 raw=0]",
        ),
        (
            r#"<?xml version="1.0"?><!--only-->"#.to_owned(),
            "NonConformant(missing document root) [decl=1 elem=0 text=0 raw=0]",
        ),
        (
            "<r/><s/>".to_owned(),
            "NonConformant(multiple roots) [decl=0 elem=1 text=0 raw=2]",
        ),
        (
            "<r></r><s></s>".to_owned(),
            "NonConformant(multiple roots) [decl=0 elem=2 text=0 raw=2]",
        ),
        (
            "<r/><!--between--><s>".to_owned(),
            "NonConformant(multiple roots) [decl=0 elem=1 text=0 raw=2]",
        ),
        (
            "<r>".to_owned(),
            "NonConformant(unterminated XML) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            "<r><a>".to_owned(),
            "NonConformant(unterminated XML) [decl=0 elem=2 text=0 raw=2]",
        ),
        ("</r>".to_owned(), "Xml [decl=0 elem=0 text=0 raw=0]"),
    ]);
}

#[test]
fn stream_accepts_a_declaration_only_as_the_first_event() {
    assert_outcomes(&[
        (
            r#"<?xml version="1.0"?><?xml version="1.0"?><r/>"#.to_owned(),
            "NonConformant(late XML declaration) [decl=1 elem=0 text=0 raw=0]",
        ),
        (
            r#"<!--first--><?xml version="1.0"?><r/>"#.to_owned(),
            "NonConformant(late XML declaration) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            " <?xml version=\"1.0\"?><r/>".to_owned(),
            "NonConformant(late XML declaration) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            r#"<r><?xml version="1.0"?></r>"#.to_owned(),
            "NonConformant(late XML declaration) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            r#"<r/><?xml version="1.0"?>"#.to_owned(),
            "NonConformant(late XML declaration) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            r#"<r></r><?xml version="1.0"?>"#.to_owned(),
            "NonConformant(late XML declaration) [decl=0 elem=2 text=0 raw=1]",
        ),
    ]);
}

#[test]
fn stream_rejects_character_content_outside_the_root() {
    assert_outcomes(&[
        (
            "text<r/>".to_owned(),
            "NonConformant(text outside document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            "<r/>text".to_owned(),
            "NonConformant(text outside document root) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            r#"<?xml version="1.0"?>text<r/>"#.to_owned(),
            "NonConformant(text outside document root) [decl=1 elem=0 text=0 raw=0]",
        ),
        (
            "<![CDATA[c]]><r/>".to_owned(),
            "NonConformant(CDATA outside document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            "<r/><![CDATA[c]]>".to_owned(),
            "NonConformant(CDATA outside document root) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            "&amp;<r/>".to_owned(),
            "NonConformant(reference outside document root) [decl=0 elem=0 text=0 raw=0]",
        ),
        (
            "<r/>&amp;".to_owned(),
            "NonConformant(reference outside document root) [decl=0 elem=1 text=0 raw=1]",
        ),
        (
            "<r></r>&amp;".to_owned(),
            "NonConformant(reference outside document root) [decl=0 elem=2 text=0 raw=1]",
        ),
    ]);
}

#[test]
fn stream_raw_flow_keeps_the_document_rules_after_a_semantic_failure() {
    let open = failing_root_open();
    let empty = failing_empty_root();
    assert_outcomes(&[
        (
            format!("{open}</r>"),
            "NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!("{open}</r>\n<!--after-->\n"),
            "NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            open.clone(),
            "NonConformant(unterminated XML) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!("{open}</r><s/>"),
            "NonConformant(multiple roots) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=4]",
        ),
        (
            format!("{open}</r><s>"),
            "NonConformant(multiple roots) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=4]",
        ),
        (
            format!("{open}</r>text"),
            "NonConformant(text outside document root) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!("{open}</r><![CDATA[c]]>"),
            "NonConformant(CDATA outside document root) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!("{open}</r>&amp;"),
            "NonConformant(reference outside document root) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!(r#"{open}</r><?xml version="1.0"?>"#),
            "NonConformant(late XML declaration) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            format!(r#"{open}<?xml version="1.0"?></r>"#),
            "NonConformant(late XML declaration) after NonConformant(unbound Requires prefix) [decl=0 elem=1 text=0 raw=3]",
        ),
        (
            empty.clone(),
            "MustUnderstand(urn:x) [decl=0 elem=0 text=0 raw=1]",
        ),
        (
            format!("{empty}<s/>"),
            "NonConformant(multiple roots) after MustUnderstand(urn:x) [decl=0 elem=0 text=0 raw=2]",
        ),
        (
            format!("{empty}text"),
            "NonConformant(text outside document root) after MustUnderstand(urn:x) [decl=0 elem=0 text=0 raw=1]",
        ),
        (
            format!(r#"{empty}<?xml version="1.0"?>"#),
            "NonConformant(late XML declaration) after MustUnderstand(urn:x) [decl=0 elem=0 text=0 raw=1]",
        ),
        (
            format!(r#"<?xml version="1.0"?>{empty}"#),
            "MustUnderstand(urn:x) [decl=1 elem=0 text=0 raw=1]",
        ),
    ]);
}

/// What the stream driver reports for `xml` under `max_events`, with a raw
/// observer that fails on every element named `fail`.
fn driven(xml: &str, max_events: usize) -> String {
    let limits = StreamLimits {
        max_events,
        ..StreamLimits::default()
    };
    let mut input = Cursor::new(xml.as_bytes());
    let result = process_markup_compatibility_stream_with_observers(
        &mut input,
        &Capabilities::new(),
        &limits,
        |element| {
            if element.name() == b"fail" {
                Err("raw")
            } else {
                Ok(())
            }
        },
        |event| match event {
            SemanticEvent::Empty(element) if element.name() == b"stop" => Err("active"),
            _ => Ok(()),
        },
    );
    let outcome = match result {
        Ok(_) => "ok".to_owned(),
        Err(StreamError::Mce {
            error,
            prior_mce_error,
            raw_error,
            active_error,
            ..
        }) => format!(
            "mce {} prior={:?} raw={raw_error:?} active={active_error:?}",
            describe(&error),
            prior_mce_error.as_ref().map(describe)
        ),
        Err(StreamError::Input {
            prior_mce_error,
            raw_error,
            active_error,
            ..
        }) => format!(
            "input prior={:?} raw={raw_error:?} active={active_error:?}",
            prior_mce_error.as_ref().map(describe)
        ),
        Err(StreamError::Callback {
            raw_error,
            active_error,
        }) => format!("callback raw={raw_error:?} active={active_error:?}"),
        Err(_) => "other".to_owned(),
    };
    outcome.replace('"', "")
}

#[test]
fn stream_driver_keeps_error_precedence_with_and_without_recovery() {
    let open = failing_root_open();
    let cases: Vec<(String, usize, &str)> = vec![
        // No semantic failure: observer failures surface at the end.
        (
            "<r><fail/></r>".to_owned(),
            100,
            "callback raw=Some(raw) active=None",
        ),
        (
            "<r><stop/><fail/></r>".to_owned(),
            100,
            "callback raw=Some(raw) active=Some(active)",
        ),
        // A later structural failure outranks the observer failures.
        (
            "<r><fail/><stop/></r><s/>".to_owned(),
            100,
            "mce NonConformant(multiple roots) prior=None raw=Some(raw) active=Some(active)",
        ),
        (
            "<r><fail/></x>".to_owned(),
            100,
            "mce Xml prior=None raw=Some(raw) active=None",
        ),
        (
            "<r><a/><a/></r>".to_owned(),
            2,
            "mce LimitExceeded { kind: Events, observed: 3, maximum: 2 } prior=None raw=None active=None",
        ),
        // Raw-only recovery: the semantic failure is primary at a clean end
        // and becomes the prior error when the raw flow fails later.
        (
            format!("{open}<fail/></r>"),
            100,
            "mce NonConformant(unbound Requires prefix) prior=None raw=Some(raw) active=None",
        ),
        (
            format!("{open}<fail/></x>"),
            100,
            "mce Xml prior=Some(NonConformant(unbound Requires prefix)) raw=Some(raw) active=None",
        ),
        (
            format!("{open}<fail/><!DOCTYPE r>"),
            100,
            "mce NonConformant(DTD and processing instructions are rejected) prior=Some(NonConformant(unbound Requires prefix)) raw=Some(raw) active=None",
        ),
        (
            format!("{open}<a/><a/><a/></r>"),
            5,
            "mce LimitExceeded { kind: Events, observed: 6, maximum: 5 } prior=Some(NonConformant(unbound Requires prefix)) raw=None active=None",
        ),
        (
            format!("{open}<fail/></r><s/>"),
            100,
            "mce NonConformant(multiple roots) prior=Some(NonConformant(unbound Requires prefix)) raw=Some(raw) active=None",
        ),
    ];
    let mismatches: Vec<String> = cases
        .iter()
        .filter_map(|(xml, max_events, expected)| {
            let found = driven(xml, *max_events);
            (found != *expected)
                .then(|| format!("{xml:?}\n   found: {found}\nexpected: {expected}"))
        })
        .collect();
    assert!(mismatches.is_empty(), "{}", mismatches.join("\n"));
}

/// The buffered processor does not check where the document ends: input that
/// mentions the MCE namespace but has no root element succeeds with whatever
/// prolog it carried, and content outside the root is dropped. The stream
/// processor rejects each of these.
#[test]
fn buffered_processor_accepts_rootless_and_stray_content_the_stream_rejects() {
    let comment_only = format!("<!--{MC}-->");
    assert_eq!(buffered(&comment_only), Ok(String::new()));
    assert_eq!(
        stream(&comment_only),
        "NonConformant(missing document root) [decl=0 elem=0 text=0 raw=0]"
    );

    let declaration_only = format!(r#"<?xml version="1.0"?><!--{MC}-->"#);
    assert_eq!(
        buffered(&declaration_only),
        Ok(r#"<?xml version="1.0"?>"#.to_owned())
    );
    assert_eq!(
        stream(&declaration_only),
        "NonConformant(missing document root) [decl=1 elem=0 text=0 raw=0]"
    );

    let late_declaration = format!(r#"<!--{MC}--><?xml version="1.0"?><r>x</r>"#);
    assert_eq!(
        buffered(&late_declaration),
        Ok(r#"<?xml version="1.0"?><r>x</r>"#.to_owned())
    );
    assert_eq!(
        stream(&late_declaration),
        "NonConformant(late XML declaration) [decl=0 elem=0 text=0 raw=0]"
    );

    let stray_text = format!(r#"before<r xmlns:mc="{MC}">x</r>after"#);
    assert_eq!(
        buffered(&stray_text),
        Ok(format!(r#"<r xmlns:mc="{MC}">x</r>"#))
    );
    assert_eq!(
        stream(&stray_text),
        "NonConformant(text outside document root) [decl=0 elem=0 text=0 raw=0]"
    );
}

/// Both processors agree on these.
#[test]
fn buffered_processor_rejects_late_declarations_multiple_roots_and_open_elements() {
    let prefix = format!("<!--{MC}-->");
    assert_eq!(
        buffered(&format!(r#"{prefix}<r/><?xml version="1.0"?>"#)),
        Err("NonConformant(late XML declaration)".to_owned())
    );
    assert_eq!(
        buffered(&format!(r#"{prefix}<r><?xml version="1.0"?></r>"#)),
        Err("NonConformant(late XML declaration)".to_owned())
    );
    assert_eq!(
        buffered(&format!("{prefix}<r/><s/>")),
        Err("NonConformant(multiple roots)".to_owned())
    );
    assert_eq!(
        buffered(&format!("{prefix}<r>")),
        Err("NonConformant(unterminated XML)".to_owned())
    );
}
