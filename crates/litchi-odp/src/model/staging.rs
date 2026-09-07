//! Auxiliary projections collected by one staging traversal.
//!
//! Errors retain the historical complete-pass order: settings, declarations,
//! then page metadata. A lower-priority error never preempts a later settings
//! error or a settings XML read failure.

use super::{declaration, page_metadata, settings};
use crate::codec::content_source::{ContentSource, Scanner as ContentSourceScanner};
use litchi_core::{Error, Result};
use quick_xml::{events::Event, reader::NsReader};

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct Metadata {
    pub(crate) settings: Option<settings::Settings>,
    pub(crate) declarations: declaration::Collection,
    pub(crate) pages: page_metadata::Collection,
}

/// The metadata projections and source-fragment skeleton collected from one
/// namespace-aware `content.xml` traversal.
#[derive(Debug)]
pub(crate) struct Staged {
    pub(crate) metadata: Metadata,
    pub(crate) content_source: Result<Option<ContentSource>>,
}

#[cfg(test)]
pub(crate) fn parse(xml: &str) -> Result<Metadata> {
    let mut settings = settings::Scanner::new(xml.len())?;
    let mut declarations = declaration::Scanner::new(xml.len());
    let mut pages = page_metadata::Scanner::new(xml.len());
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        // Settings historically owns the first complete XML traversal.
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(settings::xml_error)?;
        if matches!(event, Event::Eof) {
            break;
        }
        settings.step(&reader, &event)?;
        if let Ok(scanner) = declarations.as_mut()
            && let Err(error) = scanner.step(&reader, &event)
        {
            declarations = Err(error);
        }
        if declarations.is_ok()
            && let Ok(scanner) = pages.as_mut()
            && let Err(error) = scanner.step(&reader, &event)
        {
            pages = Err(error);
        }
        buffer.clear();
    }
    let settings = settings.finish()?;
    let declarations = declarations?.finish()?;
    let pages = pages?.finish()?;
    Ok(Metadata {
        settings,
        declarations,
        pages,
    })
}

/// Collect transaction metadata and the retained source skeleton from one
/// reader/event stream.
///
/// Metadata state machines keep their historical precedence: settings errors
/// stop immediately, declaration errors are deferred until the settings pass
/// finishes, and page errors are considered only after declarations. The
/// source-fragment scanner is auxiliary to those projections, so its structural
/// error is retained and applied only after every metadata finish step. This
/// lets the caller perform its MIME and styles setup before applying the
/// source result, matching the former staging order.
pub(crate) fn parse_with_content_source(xml: &str) -> Result<Staged> {
    let mut settings = settings::Scanner::new(xml.len())?;
    let mut declarations = declaration::Scanner::new(xml.len());
    let mut pages = page_metadata::Scanner::new(xml.len());
    let (body, has_byte_order_mark) = ContentSource::source_body(xml);
    let mut source_error = ContentSource::source_limit_error(xml);
    let mut content_source = source_error
        .is_none()
        .then(|| ContentSourceScanner::new(body));
    let mut reader = NsReader::from_str(body);
    reader.config_mut().check_end_names = true;
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    loop {
        // Settings historically owns the first complete XML traversal.
        let event = reader
            .read_event_into(&mut buffer)
            .map_err(settings::xml_error)?;
        if matches!(event, Event::Eof) {
            break;
        }
        settings.step(&reader, &event)?;
        if let Ok(scanner) = declarations.as_mut()
            && let Err(error) = scanner.step(&reader, &event)
        {
            declarations = Err(error);
        }
        if declarations.is_ok()
            && let Ok(scanner) = pages.as_mut()
            && let Err(error) = scanner.step(&reader, &event)
        {
            pages = Err(error);
        }
        if source_error.is_none() {
            let source_step = match content_source.as_mut() {
                Some(scanner) => match usize::try_from(reader.buffer_position()) {
                    Ok(end) => scanner.step(&reader, &event, end),
                    Err(_error) => Err(Error::InvalidFormat(
                        "presentation content.xml position exceeds platform limits".into(),
                    )),
                },
                None => Ok(()),
            };
            if let Err(error) = source_step {
                source_error = Some(error);
                content_source = None;
            }
        }
        buffer.clear();
    }
    let settings = settings.finish()?;
    let declarations = declarations?.finish()?;
    let pages = pages?.finish()?;
    let metadata = Metadata {
        settings,
        declarations,
        pages,
    };
    let content_source = match source_error {
        Some(error) => Err(error),
        None => content_source.map_or_else(
            || {
                Err(Error::InvalidFormat(
                    "presentation content.xml source scanner was not initialized".into(),
                ))
            },
            |scanner| scanner.finish(has_byte_order_mark),
        ),
    };
    Ok(Staged {
        metadata,
        content_source,
    })
}

#[cfg(test)]
#[allow(
    clippy::panic,
    clippy::expect_used,
    reason = "fixture differential assertions panic on a mismatched corpus"
)]
mod tests {
    use super::{Metadata, declaration, page_metadata, parse, parse_with_content_source, settings};
    use crate::codec::content_source::ContentSource;
    use litchi_core::Result;

    fn reference(xml: &str) -> Result<Metadata> {
        let settings = settings::parse_reference(xml)?;
        let declarations = declaration::parse_reference(xml)?;
        let pages = page_metadata::parse_reference(xml)?;
        Ok(Metadata {
            settings,
            declarations,
            pages,
        })
    }

    fn compare(xml: &str) {
        let normalize = |value: Result<Metadata>| value.map_err(|error| format!("{error:?}"));
        assert_eq!(
            normalize(parse(xml)),
            normalize(reference(xml)),
            "{xml:.200}"
        );
        assert_eq!(
            settings::parse(xml).map_err(|error| format!("{error:?}")),
            settings::parse_reference(xml).map_err(|error| format!("{error:?}")),
        );
        assert_eq!(
            declaration::parse(xml).map_err(|error| format!("{error:?}")),
            declaration::parse_reference(xml).map_err(|error| format!("{error:?}")),
        );
        assert_eq!(
            page_metadata::parse(xml).map_err(|error| format!("{error:?}")),
            page_metadata::parse_reference(xml).map_err(|error| format!("{error:?}")),
        );

        let normalize_fused = |value: Result<super::Staged>| {
            value
                .map(|staged| {
                    let metadata = format!("{:?}", staged.metadata);
                    let source = staged
                        .content_source
                        .map(|value| value.map(|source| format!("{source:?}")))
                        .map_err(|error| format!("{error:?}"));
                    (metadata, source)
                })
                .map_err(|error| format!("{error:?}"))
        };
        let normalize_reference = |value: Result<Metadata>| {
            value
                .map(|metadata| format!("{metadata:?}"))
                .map_err(|error| format!("{error:?}"))
        };
        let fused = normalize_fused(parse_with_content_source(xml));
        let expected_metadata = normalize_reference(reference(xml));
        let expected_source = ContentSource::parse_reference(xml)
            .map(|value| value.map(|source| format!("{source:?}")))
            .map_err(|error| format!("{error:?}"));
        match (fused, expected_metadata) {
            (Ok((metadata, source)), Ok(expected_metadata)) => {
                assert_eq!(metadata, expected_metadata, "{xml:.200}: fused metadata");
                assert_eq!(source, expected_source, "{xml:.200}: fused source");
            },
            (fused, expected) => assert_eq!(fused.map(|value| value.0), expected, "{xml:.200}"),
        }
    }

    fn wrap(body: &str) -> String {
        format!(
            "<o:document-content xmlns:o=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:p=\"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0\" xmlns:d=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" xmlns:x=\"urn:foreign\"><o:body><o:presentation>{body}</o:presentation></o:body></o:document-content>"
        )
    }

    #[test]
    fn shared_traversal_matches_independent_loops_and_namespace_scopes() {
        let fragments = [
            "",
            "<d:page d:name=\"one\"/>",
            "<d:page d:name=\"one\"><p:notes/></d:page>",
            "<p:settings p:mouse-visible=\"true\"><p:show p:name=\"custom\" p:pages=\"one,two\"/></p:settings>",
            "<p:header-decl p:name=\"header\">A &amp; B<![CDATA[<>]]></p:header-decl><d:page p:use-header-name=\"header\"/>",
            "<p:footer-decl p:name=\"footer\"/><p:date-time-decl p:name=\"date\">Today</p:date-time-decl>",
            "<x:foreign><d:page d:name=\"nested\"/></x:foreign>",
            "<d:page xmlns:d=\"urn:foreign\" d:name=\"foreign\"/><d:page d:name=\"real\"/>",
            "<page xmlns=\"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0\" name=\"unqualified\"/>",
            "<!-- comment --><d:page d:name=\"Unicode 中文 &amp; &quot;\"/>",
        ];
        for left in fragments {
            for right in fragments {
                let xml = wrap(&format!("{left}{right}"));
                compare(&xml);
                compare(&format!("\u{feff}<?xml version=\"1.0\"?>{xml}"));
                compare(
                    &xml.replace("o:", "office:")
                        .replace("xmlns:o=", "xmlns:office="),
                );
            }
        }
        for xml in ["", " ", "<", "<root>", "<root/>", "<root></wrong>"] {
            compare(xml);
        }
    }

    #[test]
    fn shared_traversal_preserves_complete_pass_error_priority() {
        let failures = [
            "<p:header-decl/>",
            "<p:header-decl p:name=\"h\"><x:child/></p:header-decl>",
            "<p:header-decl p:name=\"h\">&unknown;</p:header-decl>",
            "<d:page d:id=\"one\" xml:id=\"two\"/>",
            "<d:page xml:id=\"same\"/><d:page xml:id=\"same\"/>",
            "<p:settings p:mouse-visible=\"invalid\"/>",
            "<p:settings/><p:settings/>",
            "<p:settings><x:child/></p:settings>",
            "<p:settings><p:show p:name=\"x\" p:pages=\"one\"/><p:show p:name=\"x\" p:pages=\"two\"/></p:settings>",
            "<?active instruction?>",
            "<d:page broken=\"unterminated></d:page>",
            "<x:open></x:wrong>",
        ];
        for first in failures {
            for second in failures {
                let xml = wrap(&format!("{first}{second}"));
                assert!(reference(&xml).is_err());
                compare(&xml);
            }
        }
        let xml = wrap("<d:page d:id=\"one\" xml:id=\"two\"/><p:header-decl/>");
        let expected = declaration::parse_reference(&xml).unwrap_err().to_string();
        assert_eq!(parse(&xml).unwrap_err().to_string(), expected);
        let xml = wrap("<p:header-decl/><p:settings p:mouse-visible=\"invalid\"/>");
        let expected = settings::parse_reference(&xml).unwrap_err().to_string();
        assert_eq!(parse(&xml).unwrap_err().to_string(), expected);
    }

    #[test]
    fn shared_traversal_preserves_xml_and_page_limits() {
        compare(&" ".repeat(8 * 1024 * 1024));
        let oversized = " ".repeat(8 * 1024 * 1024 + 1);
        assert!(parse(&oversized).is_err());
        compare(&oversized);
        let pages = wrap(&"<d:page/>".repeat(65_537));
        assert!(parse(&pages).is_err());
        compare(&pages);
    }

    #[test]
    fn source_shape_errors_are_deferred_until_metadata_finishes() {
        let staged = parse_with_content_source("<root/>")
            .expect("metadata projections accept an unrelated XML root");
        assert!(staged.content_source.is_err());

        let xml = format!(
            "<root/>{}",
            wrap("<p:settings p:mouse-visible=\"invalid\"/>")
        );
        let expected = settings::parse_reference(&xml).unwrap_err().to_string();
        assert_eq!(
            parse_with_content_source(&xml).unwrap_err().to_string(),
            expected
        );

        let declaration_xml = format!("<root/>{}", wrap("<p:header-decl/>"));
        let expected = declaration::parse_reference(&declaration_xml)
            .unwrap_err()
            .to_string();
        assert_eq!(
            parse_with_content_source(&declaration_xml)
                .unwrap_err()
                .to_string(),
            expected
        );

        let page_xml = format!("<root/>{}", wrap("<d:page d:id=\"one\" xml:id=\"two\"/>"));
        let expected = page_metadata::parse_reference(&page_xml)
            .unwrap_err()
            .to_string();
        assert_eq!(
            parse_with_content_source(&page_xml)
                .unwrap_err()
                .to_string(),
            expected
        );

        let malformed_xml = format!("<root/>{}", wrap("<d:page broken=\"unterminated></d:page>"));
        let expected = reference(&malformed_xml).unwrap_err().to_string();
        assert_eq!(
            parse_with_content_source(&malformed_xml)
                .unwrap_err()
                .to_string(),
            expected
        );
    }

    #[test]
    fn fused_source_preserves_bom_prologue_and_page_spans() {
        let xml = format!(
            "\u{feff}<?xml version=\"1.0\"?>{}",
            wrap("<d:page d:name=\"bom\"/>")
        );
        let staged = parse_with_content_source(&xml).expect("BOM fixture stages");
        let source = staged
            .content_source
            .expect("BOM source result")
            .expect("BOM presentation skeleton");
        let mut prologue = String::new();
        source.write_prolog(&mut prologue);
        assert_eq!(prologue, "\u{feff}<?xml version=\"1.0\"?>");
        assert_eq!(source.page(0), Some("<d:page d:name=\"bom\"/>"));
    }

    #[test]
    fn fused_staging_matches_independent_source_reference_on_native_fixtures() {
        let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/odf/odp");
        for name in [
            "cellspan.odp",
            "text-in-image.odp",
            "tdf102223.odp",
            "tdf105502.odp",
            "tdf169979.odp",
        ] {
            let path = root.join(name);
            let presentation = crate::Presentation::open(&path)
                .unwrap_or_else(|error| panic!("{name}: opening native fixture: {error}"));
            let staged = presentation
                .staging()
                .unwrap_or_else(|error| panic!("{name}: fused staging: {error}"));
            let fused_source = staged
                .content_source
                .map(|value| value.map(|source| format!("{source:?}")))
                .map_err(|error| format!("{error:?}"));
            let reference_source = ContentSource::parse_reference(presentation.content_xml())
                .map(|value| value.map(|source| format!("{source:?}")))
                .map_err(|error| format!("{error:?}"));
            assert_eq!(fused_source, reference_source, "{name}: source projection");

            let expected_metadata = reference(presentation.content_xml())
                .unwrap_or_else(|error| panic!("{name}: metadata reference: {error}"));
            assert_eq!(
                staged.metadata, expected_metadata,
                "{name}: metadata projection"
            );
        }
    }
}
