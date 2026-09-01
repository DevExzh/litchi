//! Behavioral and fixture coverage for the MCE owner.

use super::codec::{
    active_marker_with_hash, active_offsets, find_bytes, process_markup_compatibility,
    reserve_exact,
};
use super::model::{Capabilities, Error, Limits, OffsetLimits, Report};
use std::{borrow::Cow, str};

#[cfg(test)]
mod preservation_tests {
    use super::*;

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn run_with(
        xml: &str,
        capabilities: &Capabilities,
        limits: &Limits,
    ) -> Result<(String, Report), Error> {
        let output = process_markup_compatibility(xml.as_bytes(), capabilities, limits)?;
        Ok((
            String::from_utf8(output.xml.into_owned()).expect("MCE output must remain UTF-8"),
            output.report,
        ))
    }

    fn run(xml: &str) -> Result<(String, Report), Error> {
        run_with(xml, &Capabilities::new(), &Limits::default())
    }

    #[test]
    fn preserves_exact_and_wildcard_attributes_by_expanded_name() {
        let exact = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:y="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:keep"><a y:keep="yes" y:drop="no"/></r>"#
        );
        let (xml, report) = run(&exact).unwrap();
        assert!(xml.contains(r#"y:keep="yes""#));
        assert!(!xml.contains("y:drop"));
        assert_eq!(report.preserved_attributes, 1);

        let wildcard = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:*"><a x:one="1" x:two="2"/></r>"#
        );
        let (xml, report) = run(&wildcard).unwrap();
        assert!(xml.contains(r#"x:one="1""#));
        assert!(xml.contains(r#"x:two="2""#));
        assert_eq!(report.preserved_attributes, 2);
    }

    #[test]
    fn preserves_elements_but_still_processes_their_content_and_attributes() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveElements="x:keep" mc:PreserveAttributes="x:flag"><x:keep plain="yes" x:flag="yes" x:drop="no"><x:drop/><known/></x:keep><x:discard/></r>"#
        );
        let (xml, report) = run(&source).unwrap();
        assert!(xml.contains("<x:keep"));
        assert!(xml.contains(r#"plain="yes""#));
        assert!(xml.contains(r#"x:flag="yes""#));
        assert!(xml.contains("<known"));
        assert!(!xml.contains("x:drop"));
        assert!(!xml.contains("x:discard"));
        assert_eq!(report.preserved_elements, 1);
        assert_eq!(report.preserved_attributes, 1);
    }

    #[test]
    fn redundant_ignorable_redeclaration_retains_inherited_preservation() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:*"><a mc:Ignorable="x" x:value="discarded"/></r>"#
        );
        let (xml, report) = run(&source).unwrap();
        assert!(xml.contains("x:value"));
        assert_eq!(report.preserved_attributes, 1);
    }

    #[test]
    fn understood_attributes_are_not_discarded_and_spoofed_directives_do_not_apply() {
        let understood = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x"><a x:value="kept"/></r>"#
        );
        let mut capabilities = Capabilities::new();
        capabilities.understand_namespace("urn:ext");
        let (xml, _) = run_with(&understood, &capabilities, &Limits::default()).unwrap();
        assert!(xml.contains(r#"x:value="kept""#));

        let spoofed = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:f="urn:fake" mc:Ignorable="x f" f:PreserveAttributes="x:*"><a x:value="discarded"/></r>"#
        );
        let (xml, _) = run(&spoofed).unwrap();
        assert!(!xml.contains("PreserveAttributes"));
        assert!(!xml.contains("x:value"));
    }

    #[test]
    fn rejects_invalid_preservation_tokens_and_duplicates() {
        for directive in ["keep", "missing:keep", "x:keep:extra", "x:keep x:keep"] {
            let source = format!(
                r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="{directive}"/>"#
            );
            assert!(
                run(&source).is_err(),
                "accepted invalid token list: {directive}"
            );
        }

        let wrong_namespace = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" xmlns:y="urn:other" mc:Ignorable="x" mc:PreserveElements="y:keep"/>"#
        );
        assert!(run(&wrong_namespace).is_err());
    }

    #[test]
    fn preservation_tokens_respect_the_shared_directive_bound() {
        let source = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:ext" mc:Ignorable="x" mc:PreserveAttributes="x:one x:two"/>"#
        );
        let limits = Limits {
            max_directive_tokens: 2,
            ..Limits::default()
        };
        assert!(matches!(
            run_with(&source, &Capabilities::new(), &limits),
            Err(Error::LimitExceeded(_))
        ));
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn source_offset(xml: &[u8], needle: &[u8]) -> u32 {
        let offset = find_bytes(xml, needle).expect("test element must occur in source XML");
        u32::try_from(offset).expect("test source offset must fit u32")
    }

    #[test]
    fn active_offsets_select_choice_and_fallback_in_caller_order() {
        let xml = br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:s="urn:supported" xmlns:x="urn:unsupported"><mc:AlternateContent><mc:Choice Requires="x"><inactive-choice/></mc:Choice><mc:Fallback><active-fallback/></mc:Fallback></mc:AlternateContent><mc:AlternateContent><mc:Choice Requires="s"><active-choice/></mc:Choice><mc:Fallback><inactive-fallback/></mc:Fallback></mc:AlternateContent></r>"#;
        let inactive_choice = source_offset(xml, b"<inactive-choice/>");
        let active_fallback = source_offset(xml, b"<active-fallback/>");
        let active_choice = source_offset(xml, b"<active-choice/>");
        let inactive_fallback = source_offset(xml, b"<inactive-fallback/>");
        let input = [
            active_choice,
            inactive_choice,
            inactive_choice,
            active_fallback,
            active_fallback,
            inactive_fallback,
        ];
        let mut capabilities = Capabilities::new();
        capabilities.understand_namespace("urn:supported");

        let selected =
            active_offsets(xml, &input, &capabilities, &OffsetLimits::default()).unwrap();

        assert_eq!(
            selected,
            vec![active_choice, active_fallback, active_fallback]
        );
        assert!(selected.iter().all(|offset| {
            usize::try_from(*offset)
                .ok()
                .and_then(|offset| xml.get(offset))
                == Some(&b'<')
        }));
    }

    #[test]
    fn active_offsets_fast_path_preserves_order_and_duplicates() {
        let xml = b"<r><first/><second/></r>";
        let first = source_offset(xml, b"<first/>");
        let second = source_offset(xml, b"<second/>");
        let input = [second, first, second];

        assert_eq!(
            active_offsets(xml, &input, &Capabilities::new(), &OffsetLimits::default(),).unwrap(),
            input
        );
    }

    #[test]
    fn active_offsets_reject_invalid_offsets_and_resource_limits() {
        let xml = br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><a/></r>"#;
        let outside = u32::try_from(xml.len()).unwrap();
        assert!(matches!(
            active_offsets(
                xml,
                &[outside],
                &Capabilities::new(),
                &OffsetLimits::default(),
            ),
            Err(Error::NonConformant(_))
        ));

        let offset = source_offset(xml, b"<a/>");
        let count_limited = OffsetLimits {
            max_offsets: 0,
            ..OffsetLimits::default()
        };
        assert!(matches!(
            active_offsets(xml, &[offset], &Capabilities::new(), &count_limited),
            Err(Error::LimitExceeded(_))
        ));

        let marked_limited = OffsetLimits {
            max_marked_bytes: xml.len(),
            ..OffsetLimits::default()
        };
        assert!(matches!(
            active_offsets(xml, &[offset], &Capabilities::new(), &marked_limited),
            Err(Error::LimitExceeded(_))
        ));
    }

    #[test]
    fn active_offset_markers_skip_source_collisions() {
        let hash = 0x0123_4567_89ab_cdef;
        let first = active_marker_with_hash(b"", hash).unwrap();
        let selected = active_marker_with_hash(&first, hash).unwrap();

        assert_ne!(selected, first);
        assert!(find_bytes(&first, &selected).is_none());
    }

    #[test]
    fn active_offset_allocation_error_retains_allocator_source() {
        let mut values = Vec::<u8>::new();
        let error = reserve_exact(&mut values, usize::MAX, "test active offsets").unwrap_err();

        assert!(matches!(
            &error,
            Error::Allocation {
                resource: "test active offsets",
                ..
            }
        ));
        assert!(std::error::Error::source(&error).is_some());
    }

    fn run(x: &str, c: &Capabilities) -> Result<String, Error> {
        Ok(String::from_utf8(
            process_markup_compatibility(x.as_bytes(), c, &Limits::default())?
                .xml
                .into_owned(),
        )
        .unwrap())
    }
    #[test]
    fn fast_borrowed() {
        assert!(matches!(
            process_markup_compatibility(b"<r/>", &Capabilities::new(), &Limits::default())
                .unwrap()
                .xml,
            Cow::Borrowed(_)
        ))
    }
    #[test]
    fn choice_fallback() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a="urn:a"><mc:AlternateContent><mc:Choice Requires="a"><yes/></mc:Choice><mc:Fallback><no/></mc:Fallback></mc:AlternateContent></r>"#;
        let mut c = Capabilities::new();
        assert!(run(x, &c).unwrap().contains("<no"));
        c.understand_namespace("urn:a");
        assert!(run(x, &c).unwrap().contains("<yes"))
    }
    #[test]
    fn ignore_and_unwrap() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:w"><x:no/><x:w><yes/></x:w></r>"#;
        let y = run(x, &Capabilities::new()).unwrap();
        assert!(!y.contains("<x:"));
        assert!(y.contains("<yes"))
    }

    #[test]
    fn process_content_accepts_effective_ignorable_and_namespace_wildcards() {
        let inherited = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x"><scope mc:ProcessContent="x:wrap"><x:wrap><kept/></x:wrap><x:drop><lost/></x:drop></scope></r>"#;
        let output = run(inherited, &Capabilities::new()).unwrap();
        assert!(output.contains("<kept"));
        assert!(!output.contains("<lost"));
        assert!(!output.contains("<x:wrap"));

        let wildcard = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:*"><x:first><one/></x:first><x:second><two/></x:second></r>"#;
        let output = run(wildcard, &Capabilities::new()).unwrap();
        assert!(output.contains("<one"));
        assert!(output.contains("<two"));
        assert!(!output.contains("<x:"));
    }

    #[test]
    fn redundant_ignorable_keeps_inherited_process_content() {
        let xml = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:Ignorable="x" mc:ProcessContent="x:wrap"><scope mc:Ignorable="x"><x:wrap><kept/></x:wrap></scope></r>"#;
        let output = run(xml, &Capabilities::new()).unwrap();
        assert!(output.contains("<kept"));
        assert!(!output.contains("<x:wrap"));
    }

    #[test]
    fn alternate_content_allows_ignorable_foreign_children_without_reordering_branches() {
        let xml = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" xmlns:s="urn:s" mc:Ignorable="x"><mc:AlternateContent x:meta="ok"><x:before/><mc:Choice Requires="s" x:meta="ok"><choice/></mc:Choice><mc:Fallback x:meta="ok"><fallback/></mc:Fallback><x:after/></mc:AlternateContent></r>"#;
        let output = run(xml, &Capabilities::new()).unwrap();
        assert!(output.contains("<fallback"));
        assert!(!output.contains("<choice"));
        assert!(!output.contains("x:before"));
        assert!(!output.contains("x:after"));
    }

    #[test]
    fn alternate_content_enforces_element_specific_attribute_constraints() {
        let mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        let invalid = [
            format!(
                r#"<r xmlns:mc="{mc}"><mc:AlternateContent bad="1"><mc:Choice Requires="x" xmlns:x="urn:x"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x" bad="1"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback bad="1"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice mc:Requires="x"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent xml:lang="en"><mc:Choice Requires="x"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback Requires="x"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback/><mc:Choice Requires="x"/></mc:AlternateContent></r>"#
            ),
            format!(
                r#"<r xmlns:mc="{mc}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x" xmlns:u="urn:unknown" u:bad="1"/></mc:AlternateContent></r>"#
            ),
        ];
        for source in invalid {
            assert!(
                run(&source, &Capabilities::new()).is_err(),
                "accepted invalid AlternateContent markup: {source}"
            );
        }
    }
    #[test]
    fn security_and_limits() {
        let x = r#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x="urn:x" mc:MustUnderstand="x"/>"#;
        assert!(matches!(
            run(x, &Capabilities::new()),
            Err(Error::MustUnderstand(_))
        ));
        let l = Limits {
            max_depth: 1,
            ..Limits::default()
        };
        assert!(process_markup_compatibility(b"<r xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><x/></r>",&Capabilities::new(),&l).is_err())
    }

    #[test]
    fn input_and_borrowed_output_limits_are_preflighted() {
        let xml = b"<root/>";
        let exact = Limits {
            max_input_bytes: xml.len(),
            max_output_bytes: xml.len(),
            ..Limits::default()
        };
        assert!(matches!(
            process_markup_compatibility(xml, &Capabilities::new(), &exact)
                .unwrap()
                .xml,
            Cow::Borrowed(_)
        ));

        for limits in [
            Limits {
                max_input_bytes: xml.len() - 1,
                ..exact.clone()
            },
            Limits {
                max_output_bytes: xml.len() - 1,
                ..exact.clone()
            },
        ] {
            assert!(matches!(
                process_markup_compatibility(xml, &Capabilities::new(), &limits),
                Err(Error::LimitExceeded(_))
            ));
        }
    }

    #[test]
    fn transformed_output_limit_is_exact_and_never_overcommitted() {
        let xml = br#"<r xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><a value="&amp;"/></r>"#;
        let output =
            process_markup_compatibility(xml, &Capabilities::new(), &Limits::default()).unwrap();
        let output_len = output.xml.len();
        let exact = Limits {
            max_output_bytes: output_len,
            ..Limits::default()
        };
        assert_eq!(
            process_markup_compatibility(xml, &Capabilities::new(), &exact)
                .unwrap()
                .xml
                .len(),
            output_len
        );
        let over = Limits {
            max_output_bytes: output_len - 1,
            ..Limits::default()
        };
        assert!(matches!(
            process_markup_compatibility(xml, &Capabilities::new(), &over),
            Err(Error::LimitExceeded(_))
        ));
    }

    #[test]
    fn deep_scopes_do_not_clone_effective_namespace_or_directive_sets() {
        const DEPTH: usize = 96;
        let mut xml = format!(r#"<r xmlns:mc="{}">"#, super::super::model::NAMESPACE);
        for index in 0..DEPTH {
            xml.push_str(&format!(
                r#"<e{index} xmlns:p{index}="urn:p{index}" mc:Ignorable="p{index}" mc:PreserveAttributes="p{index}:keep">"#
            ));
        }
        xml.push_str("<leaf/>");
        for index in (0..DEPTH).rev() {
            xml.push_str(&format!("</e{index}>"));
        }
        xml.push_str("</r>");

        let limits = Limits {
            max_depth: DEPTH + 2,
            max_namespace_bindings: DEPTH + 2,
            max_directive_tokens: 2,
            max_output_bytes: 8 * 1024 * 1024,
            ..Limits::default()
        };
        let output =
            process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &limits).unwrap();
        assert!(output.xml.windows(5).any(|window| window == b"<leaf"));

        let namespace_over = Limits {
            max_namespace_bindings: DEPTH + 1,
            ..limits.clone()
        };
        assert!(matches!(
            process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &namespace_over,),
            Err(Error::LimitExceeded(_))
        ));
    }

    #[test]
    fn directive_token_limit_accepts_exact_and_rejects_over() {
        let xml = format!(
            r#"<r xmlns:mc="{}" xmlns:a="urn:a" xmlns:b="urn:b" mc:Ignorable="a b"/>"#,
            super::super::model::NAMESPACE
        );
        let exact = Limits {
            max_directive_tokens: 2,
            ..Limits::default()
        };
        process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &exact).unwrap();
        let over = Limits {
            max_directive_tokens: 1,
            ..exact
        };
        assert!(matches!(
            process_markup_compatibility(xml.as_bytes(), &Capabilities::new(), &over),
            Err(Error::LimitExceeded(_))
        ));
    }

    #[test]
    fn dtd_external_identifiers_and_custom_entities_are_never_expanded() {
        let external = format!(
            r#"<!DOCTYPE r SYSTEM "https://example.invalid/entity"><r xmlns:mc="{}"/>"#,
            super::super::model::NAMESPACE
        );
        assert!(matches!(
            process_markup_compatibility(
                external.as_bytes(),
                &Capabilities::new(),
                &Limits::default(),
            ),
            Err(Error::NonConformant(_))
        ));

        let entity = format!(
            r#"<r xmlns:mc="{}">&external;</r>"#,
            super::super::model::NAMESPACE
        );
        assert!(matches!(
            process_markup_compatibility(
                entity.as_bytes(),
                &Capabilities::new(),
                &Limits::default(),
            ),
            Err(Error::NonConformant(_))
        ));
    }
}

#[cfg(test)]
mod fixture_tests {
    use super::*;
    use litchi_opc::{OpcPackage, PackURI};

    #[test]
    fn poi_styles_select_unsupported_vendor_fallbacks() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/poi/test-data/spreadsheet/style-alternate-content.xlsx"
        ))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/xl/styles.xml").unwrap())
            .unwrap();
        let output =
            process_markup_compatibility(part.blob(), &Capabilities::default(), &Limits::default())
                .unwrap();
        let xml = str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("hs:extension"));
        assert!(output.report.selected_fallbacks > 10);
    }

    #[test]
    fn libreoffice_pptx_emits_only_fallback_shape() {
        let package = OpcPackage::from_bytes(include_bytes!(
            "../../../../test-data/libreoffice-core/oox/qa/unit/data/import-mce.pptx"
        ))
        .unwrap();
        let part = package
            .get_part(&PackURI::new("/ppt/slides/slide1.xml").unwrap())
            .unwrap();
        let output =
            process_markup_compatibility(part.blob(), &Capabilities::default(), &Limits::default())
                .unwrap();
        let xml = str::from_utf8(output.xml.as_ref()).unwrap();
        assert!(!xml.contains("mc:AlternateContent"));
        assert!(!xml.contains("a14:m"));
        assert!(xml.contains("a:blipFill"));
        assert_eq!(output.report.selected_fallbacks, 1);
    }
}

#[cfg(test)]
mod streaming_0360_tests {
    use super::super::{
        Capabilities, Error, Name,
        stream::{
            RawElement, RawElementKind, SemanticEvent, StreamError, StreamLimits, StreamReport,
            process_markup_compatibility_stream,
            process_markup_compatibility_stream_with_observers,
        },
    };
    use std::{
        convert::Infallible,
        io::{self, BufRead, Cursor, Read},
    };

    const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

    fn streaming_0360_expanded(name: &Name) -> String {
        format!("{}:{}", name.namespace, name.local_name)
    }

    fn streaming_0360_event_signature(event: &SemanticEvent<'_>) -> String {
        match event {
            SemanticEvent::Start(element) => format!(
                "start:{}:{}",
                String::from_utf8_lossy(element.name()),
                streaming_0360_expanded(&element.expanded_name),
            ),
            SemanticEvent::Empty(element) => format!(
                "empty:{}:{}",
                String::from_utf8_lossy(element.name()),
                streaming_0360_expanded(&element.expanded_name),
            ),
            SemanticEvent::End(element) => format!(
                "end:{}:{}",
                String::from_utf8_lossy(element.name()),
                streaming_0360_expanded(&element.expanded_name),
            ),
            SemanticEvent::Text(text) => format!("text:{}", text.text()),
            SemanticEvent::CData(text) => format!("cdata:{}", text.text()),
            SemanticEvent::Comment(text) => format!("comment:{}", text.text()),
            SemanticEvent::Decl(decl) => {
                format!("decl:{}", String::from_utf8_lossy(decl.raw.as_ref()))
            },
            SemanticEvent::GeneralRef(reference) => {
                format!("ref:{}", String::from_utf8_lossy(reference.name.as_ref()))
            },
        }
    }

    fn streaming_0360_raw_signature(element: &RawElement<'_>) -> String {
        let kind = match element.kind {
            RawElementKind::Start => "start",
            RawElementKind::Empty => "empty",
        };
        format!("{kind}:{}", String::from_utf8_lossy(element.name()))
    }

    fn streaming_0360_active(
        xml: &[u8],
        capabilities: &Capabilities,
        limits: &StreamLimits,
    ) -> Result<(StreamReport, Vec<String>), StreamError<Infallible, &'static str>> {
        let mut input = Cursor::new(xml);
        let mut events = Vec::new();
        let report =
            process_markup_compatibility_stream(&mut input, capabilities, limits, |event| {
                events.push(streaming_0360_event_signature(&event));
                Ok::<(), &'static str>(())
            })?;
        Ok((report, events))
    }

    fn streaming_0360_both_observers(
        input: &mut dyn BufRead,
        capabilities: &Capabilities,
        limits: &StreamLimits,
    ) -> Result<(StreamReport, Vec<String>, Vec<String>), StreamError<&'static str, &'static str>>
    {
        let mut raw_events = Vec::new();
        let mut active_events = Vec::new();
        let report = process_markup_compatibility_stream_with_observers(
            input,
            capabilities,
            limits,
            |element| {
                raw_events.push(streaming_0360_raw_signature(&element));
                Ok::<(), &'static str>(())
            },
            |event| {
                active_events.push(streaming_0360_event_signature(&event));
                Ok::<(), &'static str>(())
            },
        )?;
        Ok((report, raw_events, active_events))
    }

    fn streaming_0360_callback_failure(
        input: &mut dyn BufRead,
        capabilities: &Capabilities,
        limits: &StreamLimits,
    ) -> StreamError<&'static str, &'static str> {
        process_markup_compatibility_stream_with_observers(
            input,
            capabilities,
            limits,
            |_| Err::<(), &'static str>("raw callback"),
            |_| Err::<(), &'static str>("active callback"),
        )
        .unwrap_err()
    }

    struct Streaming0360OneByte {
        bytes: Vec<u8>,
        position: usize,
    }

    impl Streaming0360OneByte {
        fn new(bytes: &[u8]) -> Self {
            Self {
                bytes: bytes.to_vec(),
                position: 0,
            }
        }
    }

    impl Read for Streaming0360OneByte {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            if output.is_empty() || self.position == self.bytes.len() {
                return Ok(0);
            }
            output[0] = self.bytes[self.position];
            self.position += 1;
            Ok(1)
        }
    }

    impl BufRead for Streaming0360OneByte {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            if self.position == self.bytes.len() {
                Ok(&[])
            } else {
                Ok(&self.bytes[self.position..self.position + 1])
            }
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.bytes.len());
        }
    }

    struct Streaming0360FailAfter {
        bytes: Vec<u8>,
        position: usize,
    }

    impl Streaming0360FailAfter {
        fn new(bytes: &[u8]) -> Self {
            Self {
                bytes: bytes.to_vec(),
                position: 0,
            }
        }
    }

    impl Read for Streaming0360FailAfter {
        fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
            let available = self.fill_buf()?;
            let amount = available.len().min(output.len());
            output[..amount].copy_from_slice(&available[..amount]);
            self.consume(amount);
            Ok(amount)
        }
    }

    impl BufRead for Streaming0360FailAfter {
        fn fill_buf(&mut self) -> io::Result<&[u8]> {
            if self.position == self.bytes.len() {
                Err(io::Error::other("streaming 0360 input failure"))
            } else {
                Ok(&self.bytes[self.position..])
            }
        }

        fn consume(&mut self, amount: usize) {
            self.position = self.position.saturating_add(amount).min(self.bytes.len());
        }
    }

    #[test]
    fn streaming_0360_projects_namespaced_events_and_counts_them() {
        let xml = br#"<p:r xmlns:p="urn:r"><p:item p:value="&amp;">text</p:item></p:r>"#;
        let (report, events) =
            streaming_0360_active(xml, &Capabilities::new(), &StreamLimits::default()).unwrap();

        assert_eq!(
            events,
            vec![
                "start:p:r:urn:r:r",
                "start:p:item:urn:r:item",
                "text:text",
                "end:p:item:urn:r:item",
                "end:p:r:urn:r:r",
            ]
        );
        assert_eq!(
            report,
            StreamReport {
                events: 5,
                ..StreamReport::default()
            }
        );
    }

    #[test]
    fn streaming_0360_raw_sees_inactive_branches_active_sees_selection() {
        let xml = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:x" xmlns:s="urn:s"><mc:AlternateContent><mc:Choice Requires="x"><x:inactive/></mc:Choice><mc:Choice Requires="s"><s:active/></mc:Choice><mc:Fallback><fallback/></mc:Fallback></mc:AlternateContent></r>"#
        );
        let mut capabilities = Capabilities::new();
        capabilities.understand_namespace("urn:s");
        let mut input = Cursor::new(xml.as_bytes());
        let (report, raw, active) =
            streaming_0360_both_observers(&mut input, &capabilities, &StreamLimits::default())
                .unwrap();

        assert!(raw.iter().any(|event| event == "start:mc:Choice"));
        assert!(raw.iter().any(|event| event == "empty:x:inactive"));
        assert!(raw.iter().any(|event| event == "empty:s:active"));
        assert_eq!(
            active,
            vec!["start:r::r", "empty:s:active:urn:s:active", "end:r::r",]
        );
        assert!(!active.iter().any(|event| event.contains("inactive")));
        assert!(!active.iter().any(|event| event.contains("fallback")));
        assert_eq!(report.alternate_content_count, 1);
        assert_eq!(report.selected_choices, 1);
        assert_eq!(report.selected_fallbacks, 0);
    }

    #[test]
    fn streaming_0360_callbacks_disable_independently_and_are_retained() {
        let xml = b"<r><one/><two/></r>";
        let mut input = Cursor::new(xml);
        let mut raw_calls = 0;
        let mut active_calls = 0;
        let error = process_markup_compatibility_stream_with_observers(
            &mut input,
            &Capabilities::new(),
            &StreamLimits::default(),
            |_| {
                raw_calls += 1;
                Err::<(), &'static str>("raw callback")
            },
            |_| {
                active_calls += 1;
                Err::<(), &'static str>("active callback")
            },
        )
        .unwrap_err();

        match error {
            StreamError::Callback {
                raw_error: Some(raw),
                active_error: Some(active),
            } => {
                assert_eq!(raw, "raw callback");
                assert_eq!(active, "active callback");
            },
            _ => panic!("expected both callback errors"),
        }
        assert_eq!(raw_calls, 1);
        assert_eq!(active_calls, 1);
    }

    #[test]
    fn streaming_0360_later_failures_keep_callback_errors_as_secondary() {
        let mut malformed = Cursor::new(b"<r></q>");
        let malformed_error = streaming_0360_callback_failure(
            &mut malformed,
            &Capabilities::new(),
            &StreamLimits::default(),
        );
        match malformed_error {
            StreamError::Mce {
                error: Error::Xml(_) | Error::NonConformant(_),
                raw_error: Some(raw),
                active_error: Some(active),
            } => {
                assert_eq!(raw, "raw callback");
                assert_eq!(active, "active callback");
            },
            _ => panic!("expected malformed XML to be primary"),
        }

        let mce = format!(r#"<r xmlns:mc="{MC}"><mc:AlternateContent/></r>"#);
        let mut mce_input = Cursor::new(mce.as_bytes());
        let mce_error = streaming_0360_callback_failure(
            &mut mce_input,
            &Capabilities::new(),
            &StreamLimits::default(),
        );
        match mce_error {
            StreamError::Mce {
                error: Error::NonConformant(_),
                raw_error: Some(raw),
                active_error: Some(active),
            } => {
                assert_eq!(raw, "raw callback");
                assert_eq!(active, "active callback");
            },
            _ => panic!("expected MCE failure to be primary"),
        }

        let mut failed_input = Streaming0360FailAfter::new(b"<r/>");
        let input_error = streaming_0360_callback_failure(
            &mut failed_input,
            &Capabilities::new(),
            &StreamLimits::default(),
        );
        match input_error {
            StreamError::Input {
                raw_error: Some(raw),
                active_error: Some(active),
                ..
            } => {
                assert_eq!(raw, "raw callback");
                assert_eq!(active, "active callback");
            },
            _ => panic!("expected input failure to be primary"),
        }
    }

    #[test]
    fn streaming_0360_event_and_count_limits_are_exact_and_typed() {
        let nested = b"<r><leaf/></r>";
        let exact_events = StreamLimits {
            max_events: 3,
            ..StreamLimits::default()
        };
        let (report, _) =
            streaming_0360_active(nested, &Capabilities::new(), &exact_events).unwrap();
        assert_eq!(report.events, 3);

        let over_events = StreamLimits {
            max_events: 2,
            ..exact_events
        };
        assert!(matches!(
            streaming_0360_active(nested, &Capabilities::new(), &over_events),
            Err(StreamError::Mce {
                error: Error::LimitExceeded(message),
                ..
            }) if message.contains("stream event count")
        ));

        let attr_xml = format!(r#"<r value="{}"/>"#, "x".repeat(64));
        let exact_attribute = StreamLimits {
            max_event_bytes: attr_xml.len(),
            ..StreamLimits::default()
        };
        streaming_0360_active(attr_xml.as_bytes(), &Capabilities::new(), &exact_attribute).unwrap();
        let over_attribute = StreamLimits {
            max_event_bytes: attr_xml.len() - 1,
            ..exact_attribute.clone()
        };
        assert!(matches!(
            streaming_0360_active(attr_xml.as_bytes(), &Capabilities::new(), &over_attribute),
            Err(StreamError::Mce {
                error: Error::LimitExceeded(message),
                ..
            }) if message.contains("stream event bytes")
        ));

        let text_xml = format!("<r>{}</r>", "t".repeat(128));
        let exact_text = StreamLimits {
            max_event_bytes: 129,
            ..StreamLimits::default()
        };
        streaming_0360_active(text_xml.as_bytes(), &Capabilities::new(), &exact_text).unwrap();
        let over_text = StreamLimits {
            max_event_bytes: 128,
            ..exact_text
        };
        assert!(matches!(
            streaming_0360_active(text_xml.as_bytes(), &Capabilities::new(), &over_text),
            Err(StreamError::Mce {
                error: Error::LimitExceeded(message),
                ..
            }) if message.contains("stream event bytes")
        ));
    }

    #[test]
    fn streaming_0360_prefix_reader_handles_split_bom_and_replays_prefix() {
        let mut bom_input = Streaming0360OneByte::new(b"\xef\xbb\xbf<r/>");
        let mut bom_events = Vec::new();
        let bom_report = process_markup_compatibility_stream(
            &mut bom_input,
            &Capabilities::new(),
            &StreamLimits::default(),
            |event| {
                bom_events.push(streaming_0360_event_signature(&event));
                Ok::<(), &'static str>(())
            },
        )
        .unwrap();
        assert_eq!(bom_report.events, 1);
        assert_eq!(bom_events, vec!["empty:r::r"]);

        let mut prefix_input = Streaming0360OneByte::new(b"<?xml version=\"1.0\"?><r/>");
        let mut prefix_events = Vec::new();
        process_markup_compatibility_stream(
            &mut prefix_input,
            &Capabilities::new(),
            &StreamLimits::default(),
            |event| {
                prefix_events.push(streaming_0360_event_signature(&event));
                Ok::<(), &'static str>(())
            },
        )
        .unwrap();
        assert_eq!(
            prefix_events,
            vec!["decl:xml version=\"1.0\"", "empty:r::r"]
        );
    }

    #[test]
    fn streaming_0360_rejects_direct_alternate_content_data() {
        for content in ["text", "<![CDATA[ ]]>"] {
            let xml = format!(
                r#"<r xmlns:mc="{MC}" xmlns:x="urn:x"><mc:AlternateContent>{content}<mc:Choice Requires="x"><yes/></mc:Choice><mc:Fallback><no/></mc:Fallback></mc:AlternateContent></r>"#
            );
            assert!(
                streaming_0360_active(
                    xml.as_bytes(),
                    &Capabilities::new(),
                    &StreamLimits::default()
                )
                .is_err(),
                "accepted direct AlternateContent data: {content}"
            );
        }
    }

    #[test]
    fn streaming_0360_rejects_custom_references_in_hidden_branches_and_prolog() {
        let hidden = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:x"><mc:AlternateContent><mc:Choice Requires="x"><x:payload>&custom;</x:payload></mc:Choice><mc:Fallback><ok/></mc:Fallback></mc:AlternateContent></r>"#
        );
        assert!(
            streaming_0360_active(
                hidden.as_bytes(),
                &Capabilities::new(),
                &StreamLimits::default()
            )
            .is_err()
        );
        assert!(
            streaming_0360_active(
                b"&custom;<r/>",
                &Capabilities::new(),
                &StreamLimits::default()
            )
            .is_err()
        );
    }

    #[test]
    fn streaming_0360_rejects_bad_document_boundaries() {
        for xml in [&b"<r></q>"[..], &b""[..], &b"<r/>tail"[..]] {
            assert!(
                streaming_0360_active(xml, &Capabilities::new(), &StreamLimits::default()).is_err(),
                "accepted malformed document: {:?}",
                xml
            );
        }
    }

    #[test]
    fn streaming_0360_expands_rebound_and_unqualified_attributes() {
        let xml =
            br#"<r xmlns:p="urn:one"><p:item xmlns:p="urn:two" p:qualified="yes" plain="ok"/></r>"#;
        let mut item = None;
        let mut input = Cursor::new(xml);
        let report = process_markup_compatibility_stream(
            &mut input,
            &Capabilities::new(),
            &StreamLimits::default(),
            |event| {
                if let SemanticEvent::Empty(element) = event
                    && element.name() == b"p:item"
                {
                    item = Some((
                        element.expanded_name.clone(),
                        element
                            .attrs()
                            .iter()
                            .map(|attribute| {
                                (
                                    attribute.expanded_name.clone(),
                                    attribute.value().to_owned(),
                                )
                            })
                            .collect::<Vec<_>>(),
                    ));
                }
                Ok::<(), &'static str>(())
            },
        )
        .unwrap();

        assert_eq!(report.events, 3);
        let (element_name, attributes) = item.expect("rebound item must be observed");
        assert_eq!(element_name.namespace, "urn:two");
        assert_eq!(element_name.local_name, "item");
        let qualified = attributes
            .iter()
            .find(|(name, _)| name.local_name == "qualified")
            .expect("qualified attribute must be observed");
        assert_eq!(qualified.0.namespace, "urn:two");
        assert_eq!(qualified.1, "yes");
        let plain = attributes
            .iter()
            .find(|(name, _)| name.local_name == "plain")
            .expect("unqualified attribute must be observed");
        assert!(plain.0.namespace.is_empty());
        assert_eq!(plain.1, "ok");
    }

    #[test]
    fn streaming_0360_processes_preserves_and_keeps_opaque_content() {
        let mut capabilities = Capabilities::new();
        capabilities.preserve_extension_element(Name {
            namespace: "urn:ext".to_owned(),
            local_name: "opaque".to_owned(),
        });
        let xml = format!(
            r#"<r xmlns:mc="{MC}" xmlns:x="urn:x" xmlns:ext="urn:ext" mc:Ignorable="x ext" mc:ProcessContent="x:unwrap" mc:PreserveElements="x:keep" mc:PreserveAttributes="x:flag"><x:unwrap><known/></x:unwrap><x:keep x:flag="yes"><x:drop/><known2/></x:keep><ext:opaque><ext:child>opaque</ext:child></ext:opaque><x:drop/></r>"#
        );
        let (report, events) =
            streaming_0360_active(xml.as_bytes(), &capabilities, &StreamLimits::default()).unwrap();

        assert!(events.iter().any(|event| event.starts_with("empty:known:")));
        assert!(
            events
                .iter()
                .any(|event| event.starts_with("start:x:keep:"))
        );
        assert!(
            events
                .iter()
                .any(|event| event.starts_with("empty:known2:"))
        );
        assert!(
            events
                .iter()
                .any(|event| event.starts_with("start:ext:opaque:"))
        );
        assert!(
            events
                .iter()
                .any(|event| event.starts_with("start:ext:child:"))
        );
        assert!(events.iter().any(|event| event == "text:opaque"));
        assert!(!events.iter().any(|event| event.contains("x:drop")));
        assert_eq!(report.unwrapped_elements, 1);
        assert_eq!(report.preserved_elements, 1);
        assert_eq!(report.preserved_attributes, 1);
        assert_eq!(report.ignored_elements, 2);
    }
}
