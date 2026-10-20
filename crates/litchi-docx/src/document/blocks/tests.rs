use super::*;
use crate::document::Snapshot;
use crate::parts::memo::VisiblePayload;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn document(body: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:w14="{W14}"><w:body>{body}</w:body></w:document>"#
    )
    .into_bytes()
}

fn body(xml: &[u8]) -> Blocks {
    let payload = VisiblePayload::project(xml).unwrap();
    let source = payload.xml(Arc::new(xml.to_vec()));
    let index = index(
        source.as_slice(),
        0,
        &[],
        ContainerRoot::Document,
        payload.positions(),
    )
    .unwrap();
    Blocks::new(source, Arc::new(index))
}

fn kinds(blocks: &Blocks) -> Vec<BlockKind> {
    blocks.iter().map(|block| block.kind()).collect()
}

fn text(block: &IndexedBlock) -> String {
    crate::paragraph::extract_word_text(block.xml_bytes()).unwrap()
}

#[test]
fn body_children_are_classified_in_visible_order_with_customxml_flattened() {
    let xml = document(concat!(
        r#"<w:bookmarkStart w:id="0" w:name="start"/>"#,
        r#"<w:p><w:r><w:t>one</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        r#"<w:customXml w:element="wrapper"><w:customXmlPr/><w:p><w:r><w:t>wrapped</w:t></w:r></w:p>"#,
        r#"<w:customXml w:element="inner"><w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl></w:customXml></w:customXml>"#,
        r#"<w:sdt><w:sdtPr><w:id w:val="-7"/><w:tag w:val="tag"/><w:alias w:val="Alias"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>control</w:t></w:r></w:p></w:sdtContent></w:sdt>"#,
        r#"<w:p><w:pPr/><m:oMathPara><m:oMath><m:r><m:t>x</m:t></m:r></m:oMath></m:oMathPara></w:p>"#,
        r#"<m:oMathPara><m:oMath/></m:oMathPara>"#,
        r#"<w:p><m:oMathPara><m:oMath/></m:oMathPara><w:r><w:t>prose</w:t></w:r></w:p>"#,
        r#"<w:p/>"#,
        r#"<w:altChunk xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId9"/>"#,
        r#"<w:sectPr/>"#,
    ));
    let blocks = body(&xml);
    assert_eq!(
        kinds(&blocks),
        [
            BlockKind::Unknown,
            BlockKind::Paragraph,
            BlockKind::Table,
            BlockKind::Paragraph,
            BlockKind::Table,
            BlockKind::ContentControl,
            BlockKind::Equation,
            BlockKind::Equation,
            BlockKind::Paragraph,
            BlockKind::Paragraph,
            BlockKind::AltChunk,
            BlockKind::SectionProperties,
        ]
    );
    assert!(blocks.raw_positions_proven());
    assert_eq!(blocks.count(BlockKind::Paragraph), 4);
    assert_eq!(blocks.count(BlockKind::Table), 2);
    assert_eq!(blocks.count(BlockKind::ContentControl), 1);
    assert_eq!(blocks.count(BlockKind::Equation), 2);

    // `/body/p[N]` counts paragraphs, wrapped ones included, equations not.
    let texts: Vec<String> = (0..blocks.count(BlockKind::Paragraph))
        .map(|ordinal| {
            blocks
                .paragraph(Position::new(ordinal))
                .unwrap()
                .text()
                .unwrap()
        })
        .collect();
    assert_eq!(texts, ["one", "wrapped", "prose", ""]);
    assert!(blocks.paragraph(Position::new(4)).is_none());
    assert_eq!(
        blocks.table(Position::new(0)).unwrap().row_count().unwrap(),
        1
    );
    assert!(blocks.table(Position::new(1)).is_some());

    // Ordinals and wrapping.
    let wrapped = blocks.nth(BlockKind::Paragraph, Position::new(1)).unwrap();
    assert!(wrapped.is_wrapped());
    assert_eq!(wrapped.ordinal().get(), 1);
    assert_eq!(wrapped.position().get(), 3);
    assert_eq!(wrapped.direct_position().map(Position::get), None);
    assert!(matches!(
        wrapped.transaction_position(),
        Err(Error::Refused(refusal)) if refusal.reason().contains("customXml")
    ));
    assert!(
        blocks
            .nth(BlockKind::Table, Position::new(1))
            .unwrap()
            .is_wrapped()
    );

    // Direct positions count every direct `w:p`, equations included, and
    // agree with the document transaction's positions.
    let snapshot = Snapshot::from_xml(xml.clone()).unwrap();
    let mut direct_paragraphs = 0;
    for block in blocks.iter() {
        let Some(paragraph) = block.paragraph() else {
            continue;
        };
        if block.is_wrapped() {
            continue;
        }
        let position = block.transaction_position().unwrap();
        assert_eq!(position.get(), direct_paragraphs);
        direct_paragraphs += 1;
        assert_eq!(
            snapshot.paragraph(position).unwrap().text().unwrap(),
            paragraph.text().unwrap()
        );
    }
    assert_eq!(direct_paragraphs, snapshot.paragraph_count());
    assert_eq!(
        blocks
            .nth(BlockKind::Table, Position::new(0))
            .unwrap()
            .transaction_position()
            .unwrap()
            .get(),
        0
    );
    assert_eq!(snapshot.table_count(), 1);
    assert_eq!(
        blocks
            .content_control(Position::new(0))
            .map(|control| control.text().unwrap()),
        Some("control".to_owned())
    );
    assert_eq!(snapshot.block_content_control_count(), 1);
    let bare_equation = blocks.nth(BlockKind::Equation, Position::new(1)).unwrap();
    assert!(bare_equation.paragraph().is_none());
    assert!(bare_equation.transaction_position().is_err());
    assert!(matches!(
        bare_equation.to_block().unwrap(),
        Block::Unknown(_)
    ));
    assert!(matches!(
        blocks
            .nth(BlockKind::Equation, Position::new(0))
            .unwrap()
            .to_block()
            .unwrap(),
        Block::Paragraph(_)
    ));
    let alt = blocks.nth(BlockKind::AltChunk, Position::new(0)).unwrap();
    let Block::Alt(chunk) = alt.to_block().unwrap() else {
        panic!("altChunk block");
    };
    assert_eq!(chunk.relationship().as_str(), "rId9");
    assert!(
        blocks
            .nth(BlockKind::SectionProperties, Position::new(0))
            .unwrap()
            .transaction_position()
            .is_err()
    );
}

#[test]
fn indexed_values_parse_in_context_and_standalone() {
    let xml = format!(
        r#"<wp:document xmlns:wp="{W}" xmlns:x="urn:x"><wp:body><wp:p x:a="1"><wp:r><wp:t>aliased</wp:t></wp:r></wp:p><wp:tbl><wp:tr><wp:tc><wp:p><wp:r><wp:t>cell</wp:t></wp:r></wp:p></wp:tc></wp:tr></wp:tbl><x:p/></wp:body></wp:document>"#
    )
    .into_bytes();
    let blocks = body(&xml);
    assert_eq!(
        kinds(&blocks),
        [BlockKind::Paragraph, BlockKind::Table, BlockKind::Unknown]
    );
    let first = blocks.get(Position::new(0)).unwrap();
    assert_eq!(first.paragraph().unwrap().text().unwrap(), "aliased");
    let standalone = first.self_contained_xml().unwrap();
    let standalone = String::from_utf8(standalone).unwrap();
    assert!(standalone.starts_with("<wp:p"));
    assert!(standalone.contains(&format!(r#"xmlns:wp="{W}""#)));
    assert!(standalone.contains(r#"xmlns:x="urn:x""#));
    let table = blocks.table(Position::new(0)).unwrap();
    assert_eq!(
        table.rows().unwrap()[0].cells().unwrap()[0].text().unwrap(),
        "cell"
    );
    assert_eq!(blocks.get(Position::new(2)).unwrap().xml_bytes(), b"<x:p/>");

    let strict = format!(
        r#"<s:document xmlns:s="{STRICT_W}"><s:body><s:p/><s:tbl/><s:sectPr/></s:body></s:document>"#
    )
    .into_bytes();
    assert_eq!(
        kinds(&body(&strict)),
        [
            BlockKind::Paragraph,
            BlockKind::Table,
            BlockKind::SectionProperties
        ]
    );
}

#[test]
fn body_index_matches_ordered_block_view_and_handles_markup_compatibility() {
    // Declared and ignorable markup only: positions are proven.
    let ignorable = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="{W14}" mc:Ignorable="w14"><w:body><w:p w14:paraId="0000002A"><w:r><w:t>a</w:t></w:r></w:p><w:p/></w:body></w:document>"#
    )
    .into_bytes();
    let blocks = body(&ignorable);
    assert!(blocks.raw_positions_proven());
    assert_eq!(
        blocks
            .nth(BlockKind::Paragraph, Position::new(1))
            .unwrap()
            .transaction_position()
            .unwrap()
            .get(),
        1
    );

    // Body-level AlternateContent selects a branch: the visible view counts
    // a paragraph the raw part does not hold directly (record 0008 W2).
    let alternate = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="{W14}"><w:body><w:p><w:r><w:t>P0</w:t></w:r></w:p><mc:AlternateContent><mc:Choice Requires="w14"><w:p><w:r><w:t>CHOICE</w:t></w:r></w:p></mc:Choice><mc:Fallback><w:p><w:r><w:t>FALLBACK</w:t></w:r></w:p></mc:Fallback></mc:AlternateContent><w:p><w:r><w:t>P2</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let blocks = body(&alternate);
    assert!(!blocks.raw_positions_proven());
    let texts: Vec<String> = blocks.iter().map(|block| text(&block)).collect();
    assert_eq!(texts, ["P0", "CHOICE", "P2"]);
    for block in blocks.iter() {
        assert!(matches!(
            block.transaction_position(),
            Err(Error::Refused(refusal)) if refusal.reason().contains("not proven")
        ));
    }
    // The raw part holds two direct paragraphs; the unchecked visible
    // position of `P2` (2) is not its raw position (1).
    assert_eq!(Snapshot::from_xml(alternate).unwrap().paragraph_count(), 2);
    assert_eq!(
        blocks
            .get(Position::new(2))
            .unwrap()
            .direct_position()
            .map(Position::get),
        Some(2)
    );
}

#[test]
fn body_index_refuses_malformed_bodies() {
    let unterminated = format!(r#"<w:document xmlns:w="{W}"><w:body><w:p>"#).into_bytes();
    assert!(
        index(
            &unterminated,
            0,
            &[],
            ContainerRoot::Document,
            RawPositions::Preserved
        )
        .is_err()
    );
    let two_bodies =
        format!(r#"<w:document xmlns:w="{W}"><w:body/><w:body><w:p/></w:body></w:document>"#)
            .into_bytes();
    assert!(matches!(
        index(&two_bodies, 0, &[], ContainerRoot::Document, RawPositions::Preserved),
        Err(Error::InvalidFormat(message)) if message.contains("multiple bodies")
    ));
    let trailing = format!(r#"<w:document xmlns:w="{W}"><w:body/></w:document>tail"#).into_bytes();
    assert!(
        index(
            &trailing,
            0,
            &[],
            ContainerRoot::Document,
            RawPositions::Preserved
        )
        .is_err()
    );
    let no_body = format!(r#"<w:document xmlns:w="{W}"/>"#).into_bytes();
    assert!(
        index(
            &no_body,
            0,
            &[],
            ContainerRoot::Document,
            RawPositions::Preserved
        )
        .unwrap()
        .entries
        .is_empty()
    );
}

#[test]
fn cell_blocks_hold_direct_children_only() {
    let xml = document(concat!(
        r#"<w:tbl><w:tr><w:tc><w:tcPr><w:tcW w:w="0"/></w:tcPr>"#,
        r#"<w:p><w:r><w:t>first</w:t></w:r></w:p>"#,
        r#"<w:tbl><w:tr><w:tc><w:p><w:r><w:t>nested</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
        r#"<w:customXml><w:p><w:r><w:t>wrapped</w:t></w:r></w:p></w:customXml>"#,
        r#"<w:p><w:r><w:t>last</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"#,
    ));
    let blocks = body(&xml);
    let cell = blocks
        .table(Position::new(0))
        .unwrap()
        .cell(0, 0)
        .unwrap()
        .unwrap();
    assert_eq!(cell.paragraphs().unwrap().len(), 4, "legacy outermost view");
    let cell_blocks = cell.blocks().unwrap();
    assert_eq!(
        kinds(&cell_blocks),
        [
            BlockKind::Paragraph,
            BlockKind::Table,
            BlockKind::Paragraph,
            BlockKind::Paragraph
        ]
    );
    let texts: Vec<String> = (0..3)
        .map(|ordinal| {
            cell_blocks
                .paragraph(Position::new(ordinal))
                .unwrap()
                .text()
                .unwrap()
        })
        .collect();
    assert_eq!(texts, ["first", "wrapped", "last"]);
    assert_eq!(
        cell_blocks
            .nth(BlockKind::Paragraph, Position::new(2))
            .unwrap()
            .direct_position()
            .map(Position::get),
        Some(1)
    );
    // A cell does not carry its document's verdict.
    assert!(!cell_blocks.raw_positions_proven());
    assert!(
        cell_blocks
            .nth(BlockKind::Paragraph, Position::new(0))
            .unwrap()
            .transaction_position()
            .is_err()
    );
    // Indexed reads require declared namespaces even for owned fragments.
    let owned = crate::Cell::new(b"<w:tc><w:tcPr/><w:p/><w:tbl/></w:tc>".to_vec());
    assert!(owned.blocks().is_err());
    assert!(
        crate::Cell::new(b"<tc><p/></tc>".to_vec())
            .blocks()
            .is_err()
    );
}

#[test]
fn content_control_identity_and_content_blocks() {
    let xml = document(concat!(
        r#"<w:sdt><w:sdtPr><w:alias w:val="Friendly"/><w:id w:val="+42"/><w:tag w:val="t&amp;g"/><w:lock w:val="sdtLocked"/></w:sdtPr><w:sdtEndPr/>"#,
        r#"<w:sdtContent><w:p><w:r><w:t>inner</w:t></w:r></w:p><w:tbl/><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:sdtContent></w:sdt>"#,
        r#"<w:sdt><w:sdtContent/></w:sdt>"#,
        r#"<w:sdt><w:sdtPr><w:id w:val="x"/></w:sdtPr></w:sdt>"#,
        r#"<w:sdt><w:sdtPr><w:tag w:val="a"/><w:tag w:val="b"/></w:sdtPr></w:sdt>"#,
    ));
    let blocks = body(&xml);
    assert_eq!(blocks.count(BlockKind::ContentControl), 4);
    let control = blocks.content_control(Position::new(0)).unwrap();
    assert_eq!(control.id().unwrap(), Some(42));
    assert_eq!(control.tag().unwrap(), Some("t&g"));
    assert_eq!(control.alias().unwrap(), Some("Friendly"));
    let content = control.blocks().unwrap();
    assert!(
        content.raw_positions_proven(),
        "inherits the body's verdict"
    );
    assert_eq!(
        kinds(&content),
        [
            BlockKind::Paragraph,
            BlockKind::Table,
            BlockKind::ContentControl
        ]
    );
    assert_eq!(
        content.paragraph(Position::new(0)).unwrap().text().unwrap(),
        "inner"
    );
    assert_eq!(
        content
            .nth(BlockKind::Paragraph, Position::new(0))
            .unwrap()
            .transaction_position()
            .unwrap()
            .get(),
        0
    );
    let nested = content.content_control(Position::new(0)).unwrap();
    assert_eq!(nested.id().unwrap(), None);
    assert_eq!(nested.blocks().unwrap().len(), 1);

    let empty = blocks.content_control(Position::new(1)).unwrap();
    assert_eq!(empty.id().unwrap(), None);
    assert!(empty.blocks().unwrap().is_empty());
    assert!(
        blocks
            .content_control(Position::new(2))
            .unwrap()
            .id()
            .is_err()
    );
    assert!(
        blocks
            .content_control(Position::new(3))
            .unwrap()
            .tag()
            .is_err()
    );

    // The ordered block view types the control and the element view keeps
    // its historical opaque value.
    let source = Arc::new(xml.clone());
    let typed =
        crate::parts::document_part::document_blocks(Arc::clone(&source), RawPositions::Preserved)
            .unwrap();
    assert!(
        matches!(&typed[0], Block::Sdt(control) if control.alias().unwrap() == Some("Friendly"))
    );
    let elements = crate::parts::document_part::document_elements(source).unwrap();
    assert!(
        matches!(&elements[0], crate::Element::Unknown(value) if value.xml_bytes().starts_with(b"<w:sdt>"))
    );
}

#[test]
fn paragraph_identifiers_read_the_start_tag_and_match_extensions() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:mc="{MC}" xmlns:w14="{W14}" mc:Ignorable="w14"><w:body><w:p w14:paraId="0000002A" w14:textId="7FFFFFFF"><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r/></w:p><w:p xmlns:v="{W14}" v:paraId="0000002B"/><w:p w14:paraId="0000002A"/><w:p><w:pPr/><m:oMathPara xmlns:m="{M}"/></w:p></w:body></w:document>"#
    )
    .into_bytes();
    let blocks = body(&xml);
    for block in blocks.iter() {
        let paragraph = block.paragraph().unwrap();
        let extensions = paragraph.extensions().unwrap();
        assert_eq!(paragraph.extension_ids().unwrap(), extensions.ids());
        assert_eq!(paragraph.para_id().unwrap(), extensions.ids().para_id());
        assert_eq!(paragraph.text_id().unwrap(), extensions.ids().text_id());
    }
    assert_eq!(
        blocks
            .paragraph(Position::new(0))
            .unwrap()
            .text_id()
            .unwrap(),
        Id::new(0x7FFF_FFFF)
    );
    let first = blocks.by_para_id(Id::new(0x2A).unwrap()).unwrap().unwrap();
    assert_eq!(
        first.position().get(),
        0,
        "a duplicated identifier resolves to its first block"
    );
    assert_eq!(
        blocks
            .by_para_id(Id::new(0x2B).unwrap())
            .unwrap()
            .unwrap()
            .position()
            .get(),
        2
    );
    assert!(blocks.by_para_id(Id::new(0x2C).unwrap()).unwrap().is_none());

    // Malformed identifiers fail like `extensions`.
    for attributes in [
        r#"w14:paraId="00000000""#,
        r#"w14:paraId="xyz""#,
        r#"w14:textId="00000001""#,
        r#"xmlns:x="urn:x" x:paraId="00000001""#,
    ] {
        let xml = format!(
            r#"<w:document xmlns:w="{W}" xmlns:w14="{W14}"><w:body><w:p {attributes}/></w:body></w:document>"#
        )
        .into_bytes();
        let blocks = body(&xml);
        let paragraph = blocks.paragraph(Position::new(0)).unwrap();
        assert!(paragraph.extensions().is_err(), "{attributes}");
        assert!(paragraph.para_id().is_err(), "{attributes}");
        assert!(
            blocks.by_para_id(Id::new(1).unwrap()).is_err(),
            "{attributes}"
        );
    }
}

#[test]
fn weight_charges_entries_kind_positions_and_identifier_table() {
    let xml = document(r#"<w:p/><w:p/><w:tbl/>"#);
    let blocks = body(&xml);
    let expected = 3 * (size_of::<BlockEntry>() + size_of::<(u32, u32)>()) + 3 * size_of::<u32>();
    assert!(blocks.index.weight_bytes() > expected);
    let large = format!(
        r#"<w:document xmlns:w="{W}" xmlns:m="{M}" xmlns:w14="{W14}" xmlns:extra="{}"><w:body><w:p/><w:p/><w:tbl/></w:body></w:document>"#,
        "x".repeat(4096)
    );
    assert!(body(large.as_bytes()).index.weight_bytes() >= blocks.index.weight_bytes() + 3 * 4096);
    assert!(size_of::<BlockEntry>() <= 48, "{}", size_of::<BlockEntry>());
}

#[test]
fn uncaptured_ranges_resolve_their_bindings_like_extensions() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:w14="{W14}"><w:body><w:p w14:paraId="0000002A"/></w:body></w:document>"#
    )
    .into_bytes();
    let start = xml.windows(4).position(|window| window == b"<w:p").unwrap();
    let length = xml[start..]
        .windows(2)
        .position(|window| window == b"/>")
        .unwrap()
        + 2;
    let paragraph = Paragraph::from_arc_range(
        Arc::new(xml.clone()),
        u32::try_from(start).unwrap(),
        u32::try_from(length).unwrap(),
    );
    assert_eq!(paragraph.para_id().unwrap(), Id::new(0x2A));
    assert_eq!(
        paragraph.extensions().unwrap().ids().para_id(),
        Id::new(0x2A)
    );
}

#[test]
fn public_values_are_send_and_sync() {
    const fn assert_shared<T: Send + Sync>() {}
    assert_shared::<Blocks>();
    assert_shared::<IndexedBlock>();
    assert_shared::<BlockContentControl>();
    assert_shared::<BlockIndex>();
}

#[test]
fn block_scan_refuses_malformed_xml_without_echoing_content() {
    for xml in [
        document(r#"<w:p secret="SENSITIVE" secret="duplicate"/>"#),
        document(r#"<undeclared:p/>"#),
        document(r#"<w:p>&SENSITIVE;</w:p>"#),
        document(r#"<w:p>&#x1;</w:p>"#),
        document(r#"<w:p x="&#xB;"/>"#),
        document("<w:p>\u{1}</w:p>"),
        document("<w:p><![CDATA[\u{1}]]></w:p>"),
        format!(r#"<!DOCTYPE document SYSTEM "SENSITIVE"><w:document xmlns:w="{W}"><w:body/></w:document>"#).into_bytes(),
        b"<foreign><body/></foreign>".to_vec(),
        format!(r#"<!--comment--><?xml version="1.0"?><w:document xmlns:w="{W}"/>"#).into_bytes(),
    ] {
        let error = index(&xml, 0, &[], ContainerRoot::Document, RawPositions::Preserved).unwrap_err();
        assert!(!error.to_string().contains("SENSITIVE"));
    }
}

#[test]
fn content_control_text_preserves_inherited_word_aliases() {
    let xml = format!(
        r#"<w:document xmlns:w="{W}" xmlns:second="{W}"><w:body><w:sdt><w:sdtContent><second:p><second:r><second:t>aliased</second:t></second:r></second:p></w:sdtContent></w:sdt></w:body></w:document>"#
    );
    let control = body(xml.as_bytes())
        .content_control(Position::new(0))
        .unwrap();
    assert_eq!(control.text().unwrap(), "aliased");
}

#[test]
fn duplicate_control_content_is_refused_by_both_read_routes() {
    let xml = document("<w:sdt><w:sdtContent/><w:sdtContent/></w:sdt>");
    let control = body(&xml).content_control(Position::new(0)).unwrap();
    assert!(control.blocks().is_err());
    assert!(control.tag().is_err());
}

#[test]
fn direct_character_content_prevents_math_only_classification() {
    for text in ["prose", "&amp;", "&#65;", "<![CDATA[prose]]>"] {
        for body_xml in [
            format!("<w:p>{text}<m:oMathPara/></w:p>"),
            format!("<w:p><m:oMathPara/>{text}</w:p>"),
        ] {
            assert_eq!(kinds(&body(&document(&body_xml))), [BlockKind::Paragraph]);
        }
    }
}
