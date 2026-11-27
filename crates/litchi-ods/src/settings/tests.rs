use super::{Settings, Snapshot};

const XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:vendor="urn:example:vendor">
  <office:body><office:spreadsheet>
    <vendor:extension vendor:flag="keep"><vendor:value>opaque</vendor:value></vendor:extension>
    <table:calculation-settings table:case-sensitive="false"/>
    <table:table table:name="Data"/>
  </office:spreadsheet></office:body>
</office:document-content>"#;

#[test]
fn no_op_commit_borrows_original_xml_and_preserves_unknown_content() {
    let snapshot = Snapshot::from_content_xml(
        XML,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let commit = snapshot
        .transaction()
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(!commit.changed());
    assert_eq!(commit.content_xml(), XML);
    assert!(commit.content_xml().contains("vendor:extension"));
}

#[test]
fn update_replaces_only_owned_element_and_keeps_unknown_siblings() {
    let snapshot = Snapshot::from_content_xml(
        XML,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.transaction();
    transaction
        .editor()
        .update(|settings| {
            settings.case_sensitive = Some(true);
            Ok(())
        })
        .expect("test fixture or operation should succeed");
    let commit = transaction
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(commit.changed());
    assert!(commit.content_xml().contains("vendor:extension"));
    assert!(
        commit
            .content_xml()
            .contains("table:case-sensitive=\"true\"")
    );
    let reparsed = Snapshot::from_content_xml(
        commit.content_xml(),
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    assert_eq!(
        reparsed
            .calculation()
            .expect("test fixture or operation should succeed")
            .case_sensitive,
        Some(true)
    );
}

#[test]
fn insertion_and_removal_are_atomic_and_contextual() {
    let source = XML.replace(
        "    <table:calculation-settings table:case-sensitive=\"false\"/>\n",
        "",
    );
    let snapshot = Snapshot::from_content_xml(
        &source,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    assert!(snapshot.calculation().is_none());

    let mut transaction = snapshot.transaction();
    transaction
        .replace(Settings {
            precision_as_shown: Some(true),
            ..Settings::default()
        })
        .expect("test fixture or operation should succeed");
    let inserted = transaction
        .commit()
        .expect("test fixture or operation should succeed")
        .into_owned();
    let inserted_snapshot = Snapshot::from_content_xml(
        &inserted,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    assert_eq!(
        inserted_snapshot
            .calculation()
            .expect("test fixture or operation should succeed")
            .precision_as_shown,
        Some(true)
    );

    let mut removal = inserted_snapshot.transaction();
    removal.editor().remove();
    let removed = removal
        .commit()
        .expect("test fixture or operation should succeed");
    assert!(removed.changed());
    assert!(!removed.content_xml().contains("calculation-settings"));
    assert!(removed.content_xml().contains("vendor:extension"));
}

#[test]
fn empty_spreadsheet_hosts_a_new_calculation_element() {
    let source = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body><o:spreadsheet/></o:body></o:document-content>"#;
    let snapshot = Snapshot::from_content_xml(
        source,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let mut transaction = snapshot.transaction();
    transaction
        .replace(Settings::default())
        .expect("test fixture or operation should succeed");
    let output = transaction
        .commit()
        .expect("test fixture or operation should succeed")
        .into_owned();
    assert!(output.contains("<o:spreadsheet><table:calculation-settings"));
    assert!(output.contains("</o:spreadsheet>"));
    assert!(
        Snapshot::from_content_xml(
            &output,
            litchi_odf_common::XmlStreamLimits::owner_default_value()
        )
        .expect("test fixture or operation should succeed")
        .calculation()
        .is_some()
    );
}

#[test]
fn invalid_typed_replacement_does_not_mutate_transaction() {
    let snapshot = Snapshot::from_content_xml(
        XML,
        litchi_odf_common::XmlStreamLimits::owner_default_value(),
    )
    .expect("test fixture or operation should succeed");
    let original = snapshot.calculation().cloned();
    let mut transaction = snapshot.transaction();
    let result = transaction.editor().update(|settings| {
        settings
            .iteration
            .get_or_insert_with(Default::default)
            .maximum_difference = Some("not-a-double".to_string());
        Ok(())
    });
    assert!(result.is_err());
    assert_eq!(transaction.calculation(), original.as_ref());
}

/// Runs `run` over the XML envelope orderings built around one root element
/// and compares each `label => outcome` line with the frozen expectation.
fn assert_envelope_matrix(
    open: &str,
    close: &str,
    empty: &str,
    expected: &[&str],
    run: impl Fn(&str) -> String,
) {
    const DECL: &str = r#"<?xml version="1.0" encoding="UTF-8"?>"#;
    let doc = format!("{open}{close}");
    let rows = [
        ("empty input", String::new()),
        ("whitespace only", " \n".to_owned()),
        ("declaration only", DECL.to_owned()),
        ("empty-element root", empty.to_owned()),
        ("start and end root", doc.clone()),
        ("declaration then root", format!("{DECL}{doc}")),
        ("declaration, whitespace, root", format!("{DECL}\n{doc}")),
        ("root then whitespace", format!("{doc}\n ")),
        ("root then comment", format!("{doc}<!--c-->")),
        ("root then PI", format!("{doc}<?pi x?>")),
        ("root then text", format!("{doc}text")),
        ("root then second root", format!("{doc}{doc}")),
        (
            "root then foreign start",
            format!("{doc}<foreign></foreign>"),
        ),
        ("root then foreign empty", format!("{doc}<foreign/>")),
        (
            "empty-element root then second root",
            format!("{empty}{doc}"),
        ),
        ("empty-element root then text", format!("{empty}text")),
        ("declaration after root", format!("{doc}{DECL}")),
        ("two declarations", format!("{DECL}{DECL}{doc}")),
        ("comment before declaration", format!("<!--c-->{DECL}{doc}")),
        ("comment before root", format!("<!--c-->{doc}")),
        ("PI before root", format!("<?pi x?>{doc}")),
        ("text before root", format!("text{doc}")),
        ("whitespace before declaration", format!(" {DECL}{doc}")),
        ("unterminated root", open.to_owned()),
        ("stray end tag", format!("{doc}</stray>")),
        ("DOCTYPE", format!("{DECL}<!DOCTYPE r>{doc}")),
        ("DOCTYPE after root", format!("{doc}<!DOCTYPE r>")),
        ("top-level CDATA", format!("<![CDATA[x]]>{doc}")),
        ("CDATA after root", format!("{doc}<![CDATA[x]]>")),
        ("entity reference before root", format!("&amp;{doc}")),
        ("entity reference after root", format!("{doc}&amp;")),
    ];
    let actual: Vec<String> = rows
        .iter()
        .map(|(label, xml)| format!("{label} => {}", run(xml)))
        .collect();
    let listing: Vec<String> = actual.iter().map(|line| format!("{line:?},")).collect();
    assert!(
        actual
            .iter()
            .map(String::as_str)
            .eq(expected.iter().copied()),
        "actual envelope matrix:\n{}",
        listing.join("\n")
    );
}

const MATRIX_ROOT: &str = concat!(
    "office:document-content ",
    "xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" ",
    "xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\"",
);
const MATRIX_BODY: &str = "<office:body><office:spreadsheet></office:spreadsheet></office:body>";
const MATRIX_CLOSE: &str = "</office:document-content>";

fn matrix_location(result: litchi_core::Result<super::codec::Location>) -> String {
    match result {
        Ok(value) => format!(
            "Ok(spreadsheet={}..{}, calculation={}, opaque={})",
            value.spreadsheet.start,
            value.spreadsheet.end,
            value.calculation.is_some(),
            value.opaque
        ),
        Err(error) => format!("{error:?}"),
    }
}

/// Feeds every event of `xml` to a [`super::codec::LocateHandler`] the way the
/// fused open parse does: a read failure or the first handler error ends the
/// pass, otherwise the handler is finished after the end of input.
fn locate_with_handler(xml: &str) -> litchi_core::Result<super::codec::Location> {
    let mut reader = quick_xml::reader::NsReader::from_str(xml);
    reader.config_mut().check_end_names = true;
    let mut buffer = Vec::new();
    let mut handler = super::codec::LocateHandler::default();
    loop {
        let before = reader.buffer_position();
        let (namespace, event) = reader
            .read_resolved_event_into(&mut buffer)
            .map_err(|error| litchi_core::Error::InvalidFormat(format!("read: {error}")))?;
        let namespace = super::codec::namespace_kind(&namespace);
        let after = reader.buffer_position();
        let end = matches!(event, quick_xml::events::Event::Eof);
        handler.on_event(
            namespace,
            &event,
            reader.resolver(),
            reader.decoder(),
            before,
            after,
        )?;
        if end {
            break;
        }
        buffer.clear();
    }
    handler.finish()
}

#[test]
fn settings_locate_envelope_matrix() {
    assert_envelope_matrix(
        &format!("<{MATRIX_ROOT}>{MATRIX_BODY}"),
        MATRIX_CLOSE,
        &format!("<{MATRIX_ROOT}/>"),
        &[
            "empty input => InvalidFormat(\"incomplete ODS content.xml document\")",
            "whitespace only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "declaration only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "empty-element root => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "start and end root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "declaration then root => Ok(spreadsheet=202..243, calculation=false, opaque=false)",
            "declaration, whitespace, root => Ok(spreadsheet=203..244, calculation=false, opaque=false)",
            "root then whitespace => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then comment => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then PI => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then text => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign start => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign empty => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then text => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "declaration after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "two declarations => Ok(spreadsheet=240..281, calculation=false, opaque=false)",
            "comment before declaration => Ok(spreadsheet=210..251, calculation=false, opaque=false)",
            "comment before root => Ok(spreadsheet=172..213, calculation=false, opaque=false)",
            "PI before root => Ok(spreadsheet=172..213, calculation=false, opaque=false)",
            "text before root => Ok(spreadsheet=168..209, calculation=false, opaque=false)",
            "whitespace before declaration => Ok(spreadsheet=203..244, calculation=false, opaque=false)",
            "unterminated root => InvalidFormat(\"incomplete ODS content.xml document\")",
            "stray end tag => InvalidFormat(\"invalid ODS content.xml: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => Ok(spreadsheet=214..255, calculation=false, opaque=false)",
            "DOCTYPE after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "top-level CDATA => Ok(spreadsheet=177..218, calculation=false, opaque=false)",
            "CDATA after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "entity reference before root => Ok(spreadsheet=169..210, calculation=false, opaque=false)",
            "entity reference after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
        ],
        |xml| matrix_location(super::codec::locate(xml)),
    );
}

#[test]
fn settings_locate_handler_envelope_matrix() {
    assert_envelope_matrix(
        &format!("<{MATRIX_ROOT}>{MATRIX_BODY}"),
        MATRIX_CLOSE,
        &format!("<{MATRIX_ROOT}/>"),
        &[
            "empty input => InvalidFormat(\"incomplete ODS content.xml document\")",
            "whitespace only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "declaration only => InvalidFormat(\"incomplete ODS content.xml document\")",
            "empty-element root => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "start and end root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "declaration then root => Ok(spreadsheet=202..243, calculation=false, opaque=false)",
            "declaration, whitespace, root => Ok(spreadsheet=203..244, calculation=false, opaque=false)",
            "root then whitespace => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then comment => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then PI => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then text => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign start => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "root then foreign empty => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then second root => InvalidFormat(\"ODS content.xml has more than one root element\")",
            "empty-element root then text => InvalidFormat(\"ODS content.xml has no office:spreadsheet host\")",
            "declaration after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "two declarations => Ok(spreadsheet=240..281, calculation=false, opaque=false)",
            "comment before declaration => Ok(spreadsheet=210..251, calculation=false, opaque=false)",
            "comment before root => Ok(spreadsheet=172..213, calculation=false, opaque=false)",
            "PI before root => Ok(spreadsheet=172..213, calculation=false, opaque=false)",
            "text before root => Ok(spreadsheet=168..209, calculation=false, opaque=false)",
            "whitespace before declaration => Ok(spreadsheet=203..244, calculation=false, opaque=false)",
            "unterminated root => InvalidFormat(\"incomplete ODS content.xml document\")",
            "stray end tag => InvalidFormat(\"read: ill-formed document: close tag `</stray>` does not match any open tag\")",
            "DOCTYPE => Ok(spreadsheet=214..255, calculation=false, opaque=false)",
            "DOCTYPE after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "top-level CDATA => Ok(spreadsheet=177..218, calculation=false, opaque=false)",
            "CDATA after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
            "entity reference before root => Ok(spreadsheet=169..210, calculation=false, opaque=false)",
            "entity reference after root => Ok(spreadsheet=164..205, calculation=false, opaque=false)",
        ],
        |xml| matrix_location(locate_with_handler(xml)),
    );
}
