use super::{Limits, RevisionType, parse_revisions_with_limits};
use crate::{Error, Paragraph};

#[test]
fn bound_revision_context_does_not_enable_legacy_attribute_prefix_fallback() {
    for attribute in [r#"w:author="unbound""#, r#"x:author="unbound""#] {
        let xml = format!(
            r#"<p xmlns="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><ins {attribute}/></p>"#
        );
        assert!(Paragraph::new(xml.into_bytes()).revisions().is_err());
    }
    let xml = br#"<w:p><w:ins xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" author="foreign" w:id="1" w:author="Alice"/></w:p>"#;
    let revisions = Paragraph::new(xml.to_vec()).revisions().unwrap();
    assert_eq!(revisions[0].author(), Some("Alice"));
}

#[test]
fn nested_revision_text_includes_namespace_bound_run_control_characters() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:x="urn:foreign"><w:ins w:id="1" w:author="A"><w:r><w:t>A</w:t><w:tab/></w:r><w:del w:id="2" w:author="B"><w:r><w:delText>B</w:delText><w:br/><w:cr></w:cr><w:noBreakHyphen/><w:softHyphen/><x:tab/></w:r></w:del></w:ins></w:p>"#;
    let revisions = Paragraph::new(xml.to_vec()).revisions().unwrap();
    assert_eq!(revisions[0].text(), "A\tB\n\n\u{2011}\u{ad}");
    assert_eq!(revisions[1].text(), "B\n\n\u{2011}\u{ad}");
}

#[test]
fn public_revision_limits_accept_exact_budgets_and_report_each_resource() {
    let xml = br#"<w:p><w:ins w:id="1" w:author="A"><w:r><w:t>x</w:t></w:r></w:ins></w:p>"#;
    let paragraph = Paragraph::new(xml.to_vec());
    let exact = Limits {
        max_source_bytes: xml.len(),
        max_events: 9,
        max_depth: 4,
        max_revisions: 1,
        max_metadata_bytes: 2,
        max_text_bytes: 1,
        max_value_bytes: 1,
        max_attributes: 2,
        max_inherited_namespaces: 0,
        max_inherited_namespace_bytes: 0,
    };
    assert_eq!(
        paragraph.revisions_with_limits(exact).unwrap()[0].text(),
        "x"
    );
    for (resource, limits) in [
        (
            "source bytes",
            Limits {
                max_source_bytes: xml.len() - 1,
                ..exact
            },
        ),
        (
            "events",
            Limits {
                max_events: 8,
                ..exact
            },
        ),
        (
            "depth",
            Limits {
                max_depth: 3,
                ..exact
            },
        ),
        (
            "records",
            Limits {
                max_revisions: 0,
                ..exact
            },
        ),
        (
            "metadata bytes",
            Limits {
                max_metadata_bytes: 1,
                ..exact
            },
        ),
        (
            "text bytes",
            Limits {
                max_text_bytes: 0,
                ..exact
            },
        ),
        (
            "value bytes",
            Limits {
                max_value_bytes: 0,
                ..exact
            },
        ),
        (
            "attributes",
            Limits {
                max_attributes: 1,
                ..exact
            },
        ),
    ] {
        assert!(
            matches!(paragraph.revisions_with_limits(limits),
            Err(Error::RevisionLimit { resource: actual, .. }) if actual == resource),
            "{resource}"
        );
    }
    assert!(
        Limits {
            max_depth: 129,
            ..exact
        }
        .validate()
        .is_err()
    );
    assert_eq!(paragraph.revisions().unwrap()[0].text(), "x");
}

#[test]
fn revision_metadata_requires_xsd_track_change_fields_and_values() {
    for xml in [
        br#"<w:p><w:ins/></w:p>"#.as_slice(),
        br#"<w:p><w:ins w:id="1"/></w:p>"#.as_slice(),
        br#"<w:p><w:ins w:author="A"/></w:p>"#.as_slice(),
        br#"<w:p><w:ins w:id="" w:author="A"/></w:p>"#.as_slice(),
        br#"<w:p><w:ins w:id="not-an-integer" w:author="A"/></w:p>"#.as_slice(),
        br#"<w:p><w:ins w:id="1" w:author="A" w:date="not-a-date"/></w:p>"#.as_slice(),
    ] {
        assert!(
            Paragraph::new(xml.to_vec()).revisions().is_err(),
            "accepted malformed revision metadata: {}",
            String::from_utf8_lossy(xml)
        );
    }

    let revisions = Paragraph::new(
        br#"<w:p><w:ins w:id=" +0007 " w:author="A" w:date="2026-01-01T00:00:00"/></w:p>"#.to_vec(),
    )
    .revisions()
    .unwrap();
    assert_eq!(revisions[0].id(), " +0007 ");
    assert_eq!(revisions[0].date(), Some("2026-01-01T00:00:00"));
}

#[test]
fn revision_parser_requires_one_root_and_rejects_top_level_content_or_late_declaration() {
    for xml in [
        br#"<w:p/><w:p/>"#.as_slice(),
        br#"text<w:p/>"#.as_slice(),
        br#"<w:p/>text"#.as_slice(),
        br#"<![CDATA[ ]]><w:p/>"#.as_slice(),
        br#"<w:p/><?xml version="1.0"?>"#.as_slice(),
        b" \n\t".as_slice(),
    ] {
        assert!(
            Paragraph::new(xml.to_vec()).revisions().is_err(),
            "accepted malformed revision envelope: {}",
            String::from_utf8_lossy(xml)
        );
    }

    let xml = b"<?xml version=\"1.0\"?> \n<w:p> \n  <w:ins w:id=\"1\" w:author=\"A\"/>\n</w:p> \n";
    assert_eq!(Paragraph::new(xml.to_vec()).revisions().unwrap().len(), 1);
}

#[test]
fn nested_revision_metadata_and_text_are_retained_in_source_order_and_charged() {
    let xml = format!(
        r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{}"><w:ins w:id="1" w:author="A" du:dateUtc="2026-01-01T00:00:00Z"><w:r><w:t>before</w:t></w:r><w:del w:id="2" w:author="B" du:dateUtc="2026-01-02T00:00:00Z"><w:r><w:delText>nested</w:delText></w:r></w:del><w:r><w:t>after</w:t></w:r></w:ins></w:p>"#,
        super::WORD_2023_DATE_UTC_NAMESPACE,
    );
    let paragraph = Paragraph::new(xml.into_bytes());
    let exact = Limits {
        max_revisions: 2,
        max_text_bytes: 23,
        ..Limits::default()
    };
    let revisions = paragraph.revisions_with_limits(exact).unwrap();
    assert_eq!(revisions.len(), 2);
    assert_eq!(
        (revisions[0].id(), revisions[0].text()),
        ("1", "beforenestedafter")
    );
    assert_eq!((revisions[1].id(), revisions[1].text()), ("2", "nested"));
    assert_eq!(revisions[1].author(), Some("B"));
    assert_eq!(revisions[1].date_utc(), Some("2026-01-02T00:00:00Z"));
    assert!(matches!(
        paragraph.revisions_with_limits(Limits {
            max_text_bytes: 22,
            ..exact
        }),
        Err(Error::RevisionLimit {
            resource: "text bytes",
            actual: 23,
            maximum: 22
        })
    ));
}

#[test]
fn paragraph_mark_and_numbering_revisions_have_metadata_without_inline_text() {
    let xml = format!(
        r#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:du="{}"><w:pPr><w:numPr><w:ins w:id="1" w:author="A"/></w:numPr><w:rPr><w:ins w:id="2" w:author="A" du:dateUtc="2026-01-01T00:00:00Z"/><w:rPrChange w:id="3" w:author="A"><w:rPr><w:del w:id="4" w:author="A" du:dateUtc="2026-01-02T00:00:00Z"/></w:rPr></w:rPrChange></w:rPr></w:pPr><w:ins w:id="5" w:author="A"><w:r><w:t>text</w:t></w:r></w:ins></w:p>"#,
        super::WORD_2023_DATE_UTC_NAMESPACE,
    );
    let revisions = Paragraph::new(xml.into_bytes()).revisions().unwrap();
    assert_eq!(
        revisions
            .iter()
            .map(|r| r.revision_type())
            .collect::<Vec<_>>(),
        vec![
            RevisionType::NumberingInsert,
            RevisionType::ParagraphMarkInsert,
            RevisionType::FormatChange,
            RevisionType::ParagraphMarkDelete,
            RevisionType::Insert,
        ]
    );
    assert!(revisions[..4].iter().all(|r| r.text().is_empty()));
    assert_eq!(revisions[4].text(), "text");
    assert_eq!(revisions[3].date_utc(), Some("2026-01-02T00:00:00Z"));
}

#[test]
fn metadata_namespace_and_decoded_byte_limits_are_contextual() {
    let xml = br#"<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:ins author="foreign" w:author="A&amp;B" w:id="1"/></w:p>"#;
    let revisions = parse_revisions_with_limits(
        xml,
        &[],
        Limits {
            max_metadata_bytes: 4,
            ..Limits::default()
        },
    )
    .unwrap();
    assert_eq!(revisions[0].author(), Some("A&B"));
    let inherited = vec![(
        Some(b"w".to_vec()),
        b"http://schemas.openxmlformats.org/wordprocessingml/2006/main".to_vec(),
    )];
    let xml = br#"<w:p><w:ins w:id="1" w:author="A"/></w:p>"#;
    assert!(matches!(
        parse_revisions_with_limits(
            xml,
            &inherited,
            Limits {
                max_inherited_namespaces: 0,
                ..Limits::default()
            }
        ),
        Err(Error::RevisionLimit {
            resource: "inherited namespaces",
            ..
        })
    ));
    let bytes = inherited[0].0.as_ref().unwrap().len() + inherited[0].1.len();
    let exact = Limits {
        max_inherited_namespace_bytes: bytes,
        ..Limits::default()
    };
    assert_eq!(
        parse_revisions_with_limits(xml, &inherited, exact)
            .unwrap()
            .len(),
        1
    );
    assert!(matches!(
        parse_revisions_with_limits(
            xml,
            &inherited,
            Limits {
                max_inherited_namespace_bytes: bytes - 1,
                ..exact
            }
        ),
        Err(Error::RevisionLimit {
            resource: "inherited namespace bytes",
            ..
        })
    ));
}

#[test]
fn empty_elements_and_metadata_markers_cannot_bypass_depth_or_content_checks() {
    let xml = br#"<w:p><w:ins w:id="1" w:author="A"/></w:p>"#;
    assert!(matches!(
        parse_revisions_with_limits(
            xml,
            &[],
            Limits {
                max_depth: 1,
                ..Limits::default()
            }
        ),
        Err(Error::RevisionLimit {
            resource: "depth",
            actual: 2,
            maximum: 1
        })
    ));
    for content in ["bad", "<![CDATA[bad]]>", "&#65;", "<w:r/>"] {
        let xml = format!(
            r#"<w:tr><w:trPr><w:ins w:id="1" w:author="A">{content}</w:ins></w:trPr></w:tr>"#
        );
        assert!(parse_revisions_with_limits(xml.as_bytes(), &[], Limits::default()).is_err());
    }
}

#[test]
fn every_table_read_facade_applies_the_record_limit() {
    let marker = br#"<w:cellIns w:id="1" w:author="A"/>"#.to_vec();
    let limits = Limits {
        max_revisions: 0,
        ..Limits::default()
    };
    assert!(matches!(
        crate::Table::new(marker.clone()).revisions_with_limits(limits),
        Err(Error::RevisionLimit {
            resource: "records",
            ..
        })
    ));
    assert!(matches!(
        crate::Row::new(marker.clone()).revisions_with_limits(limits),
        Err(Error::RevisionLimit {
            resource: "records",
            ..
        })
    ));
    assert!(matches!(
        crate::Cell::new(marker).revisions_with_limits(limits),
        Err(Error::RevisionLimit {
            resource: "records",
            ..
        })
    ));
}
