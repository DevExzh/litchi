//! Explicit opaque-control derivatives of a native Word RTF parent, not legacy layout.
use litchi_rtf::{
    Package,
    opaque::{Anchor, Context, Kind},
    write::Writer,
};

const ORIGINAL: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/word/word-rtf-minimal.rtf");
const CONTROLS: &[&[u8]] = &[
    br"\jsksu ",
    br"\jsku ",
    br"\horzvert1 ",
    br"\gcw120 ",
    br"\twoinone1 ",
    br"\nosectexpand ",
    br"\sectexpand24 ",
    br"\jclisttab ",
];
fn insert(input: &[u8], at: usize, value: &[u8]) -> Vec<u8> {
    [&input[..at], value, &input[at..]].concat()
}

#[test]
fn native_rtf_unmodeled_east_asian_controls_remain_opaque_through_body_writing() {
    let base = Package::from_bytes(ORIGINAL).unwrap();
    assert_eq!(base.opaque().len(), 4);
    let marker = b"RTF_MINIMAL_006";
    let at = ORIGINAL
        .windows(marker.len())
        .position(|v| v == marker)
        .unwrap()
        + marker.len();
    let input = insert(ORIGINAL, at, &CONTROLS.concat());
    let p = Package::from_bytes(&input).unwrap();
    assert_eq!(p.text(), base.text());
    assert_eq!(p.opaque().len(), base.opaque().len() + CONTROLS.len());
    let nodes = p
        .opaque()
        .iter()
        .filter(|n| CONTROLS.contains(&n.source()))
        .collect::<Vec<_>>();
    assert_eq!(nodes.len(), CONTROLS.len());
    for (node, bytes) in nodes.into_iter().zip(CONTROLS) {
        assert_eq!(node.kind(), Kind::ControlWord);
        assert_eq!(node.anchor(), Anchor::Body(marker.len()));
        assert_eq!(node.source(), *bytes);
    }
    assert!(p.to_bytes().unwrap() == input);
    let mut written = Vec::new();
    Writer::new(&mut written).write_document(p.raw()).unwrap();
    let reopened = Package::from_bytes(&written).unwrap();
    assert_eq!(reopened.text(), p.text());
    assert_eq!(reopened.opaque(), p.opaque());
    assert_eq!(
        reopened.raw().font_table().fonts(),
        p.raw().font_table().fonts()
    );
    let f = tempfile::NamedTempFile::with_suffix(".rtf").unwrap();
    p.save(f.path()).unwrap();
    assert!(std::fs::read(f.path()).unwrap() == input);
    let limits = litchi_rtf::SourceLimits::builder()
        .max_opaque_nodes(11)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        Package::from_bytes_with_options(
            &input,
            litchi_rtf::OpenOptions::new().source_limits(limits)
        ),
        Err(litchi_rtf::Error::Limit {
            kind: litchi_rtf::SourceLimitKind::OpaqueNodes,
            observed: 12,
            maximum: 11
        })
    ));
    assert!(base.to_bytes().unwrap() == ORIGINAL);
}

#[test]
fn native_rtf_legacy_control_header_derivative_refuses_lossy_canonical_publication() {
    let base = Package::from_bytes(ORIGINAL).unwrap();
    assert!(ORIGINAL.starts_with(br"{\rtf1"));
    let header = [
        br"{\header\pard ".as_slice(),
        &CONTROLS.concat(),
        b"INERT_HEADER_062}",
    ]
    .concat();
    let input = insert(ORIGINAL, 6, &header);
    let p = Package::from_bytes(&input).unwrap();
    assert_eq!(p.text(), base.text());
    assert_eq!(p.opaque().len(), base.opaque().len() + CONTROLS.len());
    let nodes = p
        .opaque()
        .iter()
        .filter(|n| CONTROLS.contains(&n.source()))
        .collect::<Vec<_>>();
    assert_eq!(nodes.len(), CONTROLS.len());
    for (node, bytes) in nodes.into_iter().zip(CONTROLS) {
        assert_eq!(node.source(), *bytes);
        assert!(matches!(
            node.anchor(),
            Anchor::Structural {
                context: Context::HeaderFooter,
                ..
            }
        ));
    }
    assert!(p.to_bytes().unwrap() == input);
    let mut output = Vec::new();
    let error = Writer::new(&mut output)
        .write_document(p.raw())
        .unwrap_err();
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    assert_eq!(
        error.to_string(),
        "RTF canonical rewrite cannot preserve a structurally scoped opaque node"
    );
    assert!(output.is_empty());
    assert!(
        Package::from_bytes(&p.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
            == input
    );
    assert!(base.to_bytes().unwrap() == ORIGINAL);
}
