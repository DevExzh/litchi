//! Expanded-name duplicate-attribute coverage for the MCE processors.
//!
//! XML namespace declarations apply to the whole start tag, and XML attribute
//! uniqueness is defined by the expanded `(namespace, local-name)` pair.  The
//! tests here keep that rule visible at the MCE boundaries where attributes
//! can otherwise be discarded: compatibility directives, opaque extensions,
//! skipped elements, and unselected AlternateContent branches.

use std::{borrow::Cow, convert::Infallible, io::Cursor};

use super::stream::{StreamError, StreamLimits, StreamReport, process_markup_compatibility_stream};
use super::{
    Capabilities, Error, Limits, NAMESPACE, Name, OffsetLimits, active_offsets,
    process_markup_compatibility,
};

const XML_NAMESPACE: &str = "http://www.w3.org/XML/1998/namespace";

fn tree_result(xml: &str) -> Result<Vec<u8>, Error> {
    process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &Limits::default())
        .map(|output| output.xml.into_owned())
}

fn stream_result(xml: &str) -> Result<StreamReport, StreamError<Infallible, Infallible>> {
    let mut input = Cursor::new(xml.as_bytes());
    process_markup_compatibility_stream(
        &mut input,
        &Capabilities::new(),
        &StreamLimits::default(),
        |_| Ok::<(), Infallible>(()),
    )
}

fn assert_tree_and_stream_duplicate(xml: &str) {
    match tree_result(xml) {
        Err(Error::NonConformant(message)) => assert_eq!(message, "duplicate attribute", "{xml}"),
        other => panic!("expected a typed tree duplicate-attribute error, got {other:?}: {xml}"),
    }

    match stream_result(xml) {
        Err(StreamError::Mce {
            error: Error::NonConformant(message),
            ..
        }) => assert_eq!(message, "duplicate attribute", "{xml}"),
        other => {
            panic!("expected a typed stream duplicate-attribute error, got {other:?}: {xml}")
        },
    }
}

fn assert_tree_and_stream_accept(xml: &str) {
    tree_result(xml)
        .unwrap_or_else(|error| panic!("tree rejected legal attributes: {error}: {xml}"));
    stream_result(xml)
        .unwrap_or_else(|error| panic!("stream rejected legal attributes: {error:?}: {xml}"));
}

#[test]
fn expanded_duplicates_cover_aliases_and_late_namespace_declarations() {
    let cases = vec![
        // Two prefixes name the same ordinary namespace and local part.
        r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:shared" xmlns:b="urn:shared" a:k="one" b:k="two"/>"#.to_owned(),
        // `xml` has a fixed namespace identity regardless of the alias used.
        format!(
            r#"<r xmlns:mc="{}" xmlns:x="{}" x:lang="one" xml:lang="two"/>"#,
            NAMESPACE,
            XML_NAMESPACE
        ),
        // MCE directive attributes are ordinary expanded XML attributes too.
        r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:m="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:ignored" mc:Ignorable="x" m:Ignorable="x"/>"#.to_owned(),
        // Namespace declarations apply even when they appear after the use.
        r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:shared" a:k="one" b:k="two" xmlns:b="urn:shared"/>"#.to_owned(),
    ];

    for xml in &cases {
        assert_tree_and_stream_duplicate(xml);
    }
}

#[test]
fn expanded_duplicates_are_checked_in_opaque_descendants() {
    let mut capabilities = Capabilities::new();
    capabilities.preserve_extension_element(Name {
        namespace: "urn:extension".to_owned(),
        local_name: "opaque".to_owned(),
    });
    let xml = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:e="urn:extension" xmlns:a="urn:shared" xmlns:b="urn:shared"><e:opaque><child a:k="one" b:k="two"/></e:opaque></r>"#;

    let tree = process_markup_compatibility(xml.as_bytes(), &capabilities, &Limits::default());
    match tree {
        Err(Error::NonConformant(message)) => assert_eq!(message, "duplicate attribute"),
        other => panic!("opaque descendant was not validated by the tree: {other:?}"),
    }

    let mut input = Cursor::new(xml.as_bytes());
    let stream = super::stream::process_markup_compatibility_stream(
        &mut input,
        &capabilities,
        &StreamLimits::default(),
        |_| Ok::<(), Infallible>(()),
    );
    assert!(matches!(
        stream,
        Err(StreamError::Mce {
            error: Error::NonConformant(message),
            ..
        }) if message == "duplicate attribute"
    ));
}

#[test]
fn expanded_duplicates_are_checked_in_skipped_and_unselected_branches() {
    let skipped = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:skipped" xmlns:a="urn:shared" xmlns:b="urn:shared" mc:Ignorable="x"><x:skip><child a:k="one" b:k="two"/></x:skip><tail/></r>"#;
    assert_tree_and_stream_duplicate(skipped);

    let unselected = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:u="urn:unsupported" xmlns:a="urn:shared" xmlns:b="urn:shared"><mc:AlternateContent><mc:Choice Requires="u"><branch a:k="one" b:k="two"/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent></r>"#;
    assert_tree_and_stream_duplicate(unselected);
}

fn attribute_list(count: usize, distinct_namespace: bool, distinct_local: bool) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{}""#, NAMESPACE);
    for index in 0..count {
        let namespace = if distinct_namespace {
            format!("urn:namespace-{index}")
        } else {
            "urn:shared".to_owned()
        };
        xml.push_str(&format!(r#" xmlns:p{index}="{namespace}""#));
    }
    for index in 0..count {
        let local = if distinct_local {
            format!("key{index}")
        } else {
            "key".to_owned()
        };
        xml.push_str(&format!(r#" p{index}:{local}="{index}""#));
    }
    xml.push_str("/>");
    xml
}

fn duplicate_attribute_list(count: usize, first: usize, second: usize) -> String {
    let mut xml = format!(r#"<r xmlns:mc="{}""#, NAMESPACE);
    for index in 0..count {
        xml.push_str(&format!(r#" xmlns:p{index}="urn:shared""#));
    }
    for index in 0..count {
        let local = if index == second {
            format!("key{first}")
        } else {
            format!("key{index}")
        };
        xml.push_str(&format!(r#" p{index}:{local}="{index}""#));
    }
    xml.push_str("/>");
    xml
}

#[test]
fn expanded_duplicate_validation_covers_inline_and_sorted_attribute_paths() {
    // The 9th name spills past the eight-entry inline scratch; the last case
    // collides only among overflow entries, after the inline names are gone.
    for (count, first, second) in [(8, 0, 7), (9, 0, 8), (33, 16, 32)] {
        let duplicate = duplicate_attribute_list(count, first, second);
        assert_tree_and_stream_duplicate(&duplicate);
    }
}

#[test]
fn expanded_duplicate_validation_keeps_distinct_names_legal_at_each_size() {
    for count in [8, 9, 33] {
        // One namespace with distinct local names is legal.
        assert_tree_and_stream_accept(&attribute_list(count, false, true));
        // One local name in distinct namespaces is also legal.
        assert_tree_and_stream_accept(&attribute_list(count, true, false));
    }
}

#[test]
fn active_offsets_rejects_the_same_expanded_duplicate_input() {
    let xml = br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:shared" xmlns:b="urn:shared" a:k="one" b:k="two"/>"#;
    let offset = u32::try_from(
        xml.iter()
            .position(|byte| *byte == b'<')
            .expect("root start"),
    )
    .expect("source offset fits");

    let result = active_offsets(
        xml,
        &[offset],
        &Capabilities::new(),
        &OffsetLimits::default(),
    );
    assert!(matches!(
        result,
        Err(Error::NonConformant(message)) if message == "duplicate attribute"
    ));
}

#[test]
fn no_mce_inputs_remain_borrowed_and_byte_exact() {
    // The legacy processor intentionally has a no-MCE fast path.  Expanded
    // duplicate validation belongs to MCE processing and must not turn this
    // unrelated input into an owned or rewritten buffer.
    let source = br#"<r xmlns:a="urn:shared" xmlns:b="urn:shared" a:k="one" b:k="two"/>"#;
    let output = process_markup_compatibility(source, &Capabilities::new(), &Limits::default())
        .expect("no-MCE input uses the borrowed fast path");
    assert!(matches!(&output.xml, Cow::Borrowed(bytes) if *bytes == source));
    assert_eq!(output.xml.as_ref(), source);
}
