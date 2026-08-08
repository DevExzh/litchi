use litchi_rtf::opaque::{Anchor, Context, Kind};
use litchi_rtf::{
    DefaultTabWidthPolicy, Document, ParseLimits, RtfDocument, RtfError, RtfWriter, WriterOptions,
};
use std::io::{self, Write};

#[test]
fn immutable_snapshot_preserves_unknown_syntax_byte_for_byte() {
    let source = br"{\rtf1\ansi A\future42 {\*\vendor\page1 inert}B}";
    let document = Document::from_bytes(source).unwrap();

    assert_eq!(document.text(), "AB");
    assert_eq!(document.opaque().len(), 2);
    assert_eq!(document.opaque()[0].kind(), Kind::ControlWord);
    assert_eq!(document.opaque()[0].anchor(), Anchor::Body(1));
    assert_eq!(document.opaque()[1].kind(), Kind::Destination);
    assert_eq!(document.opaque()[1].anchor(), Anchor::Body(1));
    assert_eq!(document.to_bytes().unwrap(), source);
}

#[test]
fn canonical_writer_reinserts_inert_nodes_at_their_body_anchor() {
    let source = br"{\rtf1\ansi A\future42 {\*\vendor opaque}B}";
    let document = RtfDocument::parse_bytes(source).unwrap();
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();

    let serialized = String::from_utf8(output.clone()).unwrap();
    let control = serialized.find("\\future42 ").unwrap();
    let destination = serialized.find("{\\*\\vendor opaque}").unwrap();
    assert!(control < destination);
    assert_eq!(RtfDocument::parse_bytes(&output).unwrap().text(), "AB");
}

#[test]
fn opaque_limits_are_typed_and_exact() {
    let control = r"{\rtf1\future1 x}";
    let limits = ParseLimits::default().with_max_opaque_nodes(0);
    assert!(matches!(
        RtfDocument::parse_with_limits(control, limits),
        Err(RtfError::LimitExceeded {
            resource: "opaque nodes",
            observed: 1,
            limit: 0,
        })
    ));

    let destination = r"{\rtf1{\*\vendor 1234}x}";
    let node_bytes = r"{\*\vendor 1234}".len();
    let limits = ParseLimits::default().with_max_opaque_node_bytes(node_bytes - 1);
    assert!(matches!(
        RtfDocument::parse_with_limits(destination, limits),
        Err(RtfError::LimitExceeded {
            resource: "opaque node bytes",
            observed,
            limit,
        }) if observed == node_bytes && limit == node_bytes - 1
    ));

    let limits = ParseLimits::default().with_max_total_opaque_bytes(node_bytes - 1);
    assert!(matches!(
        RtfDocument::parse_with_limits(destination, limits),
        Err(RtfError::LimitExceeded {
            resource: "opaque bytes",
            observed,
            limit,
        }) if observed == node_bytes && limit == node_bytes - 1
    ));
}

#[test]
fn malformed_unknown_destination_is_not_silently_preserved() {
    assert!(RtfDocument::parse(r"{\rtf1{\*\vendor unterminated}").is_err());
}

#[test]
fn nested_header_syntax_keeps_its_owner_and_refuses_reparenting() {
    let source = br"{\rtf1{\header H\future7 {\*\vendor inert}I}Body}";
    let snapshot = Document::from_bytes(source).unwrap();

    assert_eq!(snapshot.to_bytes().unwrap(), source);
    assert_eq!(snapshot.opaque().len(), 2);
    for node in snapshot.opaque() {
        assert!(matches!(
            node.anchor(),
            Anchor::Structural {
                context: Context::HeaderFooter,
                token: _,
                depth: _,
            }
        ));
    }
    let first = match snapshot.opaque()[0].anchor() {
        Anchor::Structural { token, .. } => token,
        Anchor::Body(_) => unreachable!(),
    };
    let second = match snapshot.opaque()[1].anchor() {
        Anchor::Structural { token, .. } => token,
        Anchor::Body(_) => unreachable!(),
    };
    assert!(first < second);

    let document = RtfDocument::parse_bytes(source).unwrap();
    let mut output = Vec::new();
    let error = RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap_err();
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    assert!(output.is_empty());
}

#[test]
fn nested_table_syntax_keeps_table_context_and_refuses_reparenting() {
    let source = br"{\rtf1\trowd\cellx2000\intbl A\future7 {\*\vendor inert}B\cell\row}";
    let snapshot = Document::from_bytes(source).unwrap();

    assert_eq!(snapshot.to_bytes().unwrap(), source);
    assert_eq!(snapshot.opaque().len(), 2);
    assert!(snapshot.opaque().iter().all(|node| matches!(
        node.anchor(),
        Anchor::Structural {
            context: Context::Table,
            ..
        }
    )));

    let document = RtfDocument::parse_bytes(source).unwrap();
    let mut output = Vec::new();
    assert!(
        RtfWriter::new(&mut output)
            .write_document(&document)
            .is_err()
    );
    assert!(output.is_empty());
}

#[test]
fn field_and_note_syntax_retain_their_structural_contexts() {
    for (source, context) in [
        (
            br"{\rtf1 A{\field{\*\fldinst TEST \future7 {\*\vendor inert}}{\fldrslt X}}B}"
                .as_slice(),
            Context::Field,
        ),
        (
            br"{\rtf1 A{\footnote N\future7 {\*\vendor inert}M}B}".as_slice(),
            Context::Note,
        ),
    ] {
        let snapshot = Document::from_bytes(source).unwrap();
        assert_eq!(snapshot.to_bytes().unwrap(), source);
        assert_eq!(snapshot.opaque().len(), 2);
        assert!(snapshot.opaque().iter().all(|node| matches!(
            node.anchor(),
            Anchor::Structural {
                context: actual,
                ..
            } if actual == context
        )));

        let document = RtfDocument::parse_bytes(source).unwrap();
        let mut output = Vec::new();
        assert!(
            RtfWriter::new(&mut output)
                .write_document(&document)
                .is_err()
        );
        assert!(output.is_empty());
    }
}

#[derive(Default)]
struct WriteProbe {
    touched: bool,
}

impl Write for WriteProbe {
    fn write(&mut self, _buffer: &[u8]) -> io::Result<usize> {
        self.touched = true;
        Err(io::Error::other("unexpected sink write"))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn default_options_are_the_only_exact_source_passthrough() {
    let structural = br"{\rtf1{\header H{\*\vendor inert}}Body}";
    let snapshot = Document::from_bytes(structural).unwrap();

    let mut exact = Vec::new();
    RtfWriter::new(&mut exact).write(&snapshot).unwrap();
    assert_eq!(exact, structural);

    let options = WriterOptions {
        default_tab_width: DefaultTabWidthPolicy::Override(960),
        ..WriterOptions::default()
    };
    let mut probe = WriteProbe::default();
    let error = RtfWriter::with_options(&mut probe, options)
        .write(&snapshot)
        .unwrap_err();
    assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
    assert!(!probe.touched);
}

#[test]
fn nondefault_options_canonicalize_representable_opaque_syntax() {
    let source = br"{\rtf1 A\future7 B}";
    let snapshot = Document::from_bytes(source).unwrap();
    let options = WriterOptions {
        default_font: 3,
        ..WriterOptions::default()
    };
    let mut output = Vec::new();
    RtfWriter::with_options(&mut output, options)
        .write(&snapshot)
        .unwrap();

    assert_ne!(output, source);
    let text = String::from_utf8(output).unwrap();
    assert!(text.contains("\\deff3"));
    assert!(text.contains("\\future7 "));
}

#[test]
fn default_snapshot_is_exact_without_opaque_syntax() {
    let source = br"{\rtf1\ansi   Plain text}";
    let snapshot = Document::from_bytes(source).unwrap();

    assert!(snapshot.opaque().is_empty());
    assert_eq!(snapshot.to_bytes().unwrap(), source);
}

#[test]
fn font_and_color_extensions_are_preserved_or_refused_atomically() {
    for source in [
        br"{\rtf1{\fonttbl{\f0\fnil Arial;{\*\vendorfont inert}}}\f0 Body}".as_slice(),
        br"{\rtf1{\colortbl;{\*\vendorcolor inert}\red1\green2\blue3;}Body}".as_slice(),
    ] {
        let snapshot = Document::from_bytes(source).unwrap();
        assert_eq!(snapshot.to_bytes().unwrap(), source);
        assert_eq!(snapshot.opaque().len(), 1);
        assert!(matches!(
            snapshot.opaque()[0].anchor(),
            Anchor::Structural {
                context: Context::Metadata,
                ..
            }
        ));

        let options = WriterOptions {
            default_tab_width: DefaultTabWidthPolicy::Override(960),
            ..WriterOptions::default()
        };
        let mut probe = WriteProbe::default();
        let error = RtfWriter::with_options(&mut probe, options)
            .write(&snapshot)
            .unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
        assert!(!probe.touched);
    }
}

#[test]
fn style_and_list_extensions_are_preserved_or_refused_atomically() {
    for source in [
        br"{\rtf1{\stylesheet{\s0 Normal;{\*\vendorstyle inert}}}\s0 Body}".as_slice(),
        br"{\rtf1{\*\listtable{\*\vendorlist inert}}Body}".as_slice(),
    ] {
        let snapshot = Document::from_bytes(source).unwrap();
        assert_eq!(snapshot.to_bytes().unwrap(), source);
        assert_eq!(snapshot.opaque().len(), 1);
        assert!(matches!(
            snapshot.opaque()[0].anchor(),
            Anchor::Structural {
                context: Context::Metadata,
                ..
            }
        ));

        let options = WriterOptions {
            default_tab_width: DefaultTabWidthPolicy::Override(960),
            ..WriterOptions::default()
        };
        let mut probe = WriteProbe::default();
        let error = RtfWriter::with_options(&mut probe, options)
            .write(&snapshot)
            .unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
        assert!(!probe.touched);
    }
}
