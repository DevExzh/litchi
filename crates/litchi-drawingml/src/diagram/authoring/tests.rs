#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::indexing_slicing,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::diagram::Limits;
use crate::diagram::edit::{self, Outline};

fn sample(layout: Layout) -> Spec {
    Spec::parse_outline(
        layout,
        "Plan\n\tScope\n\tBudget\nBuild & test\nShip <now>",
        &Limits::default(),
    )
    .unwrap()
}

#[test]
fn outline_round_trips_and_nests() {
    let spec = sample(Layout::Process);
    assert_eq!(spec.nodes().len(), 3);
    assert_eq!(spec.nodes()[0].children.len(), 2);
    assert_eq!(
        spec.outline(&Limits::default()).unwrap(),
        "Plan\n\tScope\n\tBudget\nBuild & test\nShip <now>"
    );
    let limits = Limits::default();
    let spaced = Spec::parse_outline(Layout::List, "A\n  B\n    C\\nD\\tE", &limits).unwrap();
    assert_eq!(spaced.outline(&limits).unwrap(), "A\n\tB\n\t\tC\nD\tE");
    assert!(Spec::parse_outline(Layout::List, "A\n\t\tB", &limits).is_err());
    assert!(Spec::parse_outline(Layout::List, "\tA", &limits).is_err());
}

#[test]
fn outline_applies_custom_limits_before_building_the_tree() {
    let defaults = Limits::default();
    let limits = |nodes, node_depth, node_text_bytes| {
        Limits::new(
            defaults.input_bytes_value(),
            defaults.output_bytes_value(),
            defaults.xml_nodes_value(),
            defaults.xml_depth_value(),
            defaults.points_value(),
            defaults.connections_value(),
            defaults.data_text_bytes_value(),
            nodes,
            node_depth,
            node_text_bytes,
        )
    };
    assert!(matches!(
        Spec::parse_outline(Layout::List, "A\nB", &limits(1, 7, 32 * 1024)),
        Err(Error::Limit {
            kind: crate::LimitKind::SmartArtNodes,
            ..
        })
    ));
    assert!(matches!(
        Spec::parse_outline(Layout::List, "A\n\tB\n\t\tC", &limits(100, 1, 32 * 1024)),
        Err(Error::Limit {
            kind: crate::LimitKind::SmartArtNodeDepth,
            ..
        })
    ));
    assert!(matches!(
        Spec::parse_outline(Layout::List, "AB", &limits(100, 7, 1)),
        Err(Error::Limit {
            kind: crate::LimitKind::SmartArtNodeTextBytes,
            ..
        })
    ));
}

#[test]
fn every_layout_writes_parts_that_read_back() {
    for layout in Layout::ALL {
        let limits = Limits::default();
        let spec = sample(layout);
        let data = spec.data_xml(Some("rId7"), &limits).unwrap();
        let outline = Outline::read(data.as_bytes(), &limits).unwrap();
        assert_eq!(
            outline.layout_type_id.as_deref(),
            Some(layout.unique_id().as_str())
        );
        assert_eq!(outline.drawing_relationship.as_deref(), Some("rId7"));
        let texts: Vec<_> = outline.entries.iter().map(|e| e.text.as_str()).collect();
        assert_eq!(
            texts,
            ["Plan", "Scope", "Budget", "Build & test", "Ship <now>"]
        );
        assert_eq!(outline.entries[1].parent, Some(0));
        assert_eq!(outline.entries[1].depth, 1);
        assert_eq!(Spec::from_outline(layout, &outline), spec);
        // The legacy validating reader accepts the model.
        crate::diagram::DiagramDataModel::parse(&data, &limits).unwrap();
        let layout_xml = spec.layout_xml(&limits).unwrap();
        crate::diagram::DiagramDefinition::parse_layout(&layout_xml, &limits).unwrap();
        crate::diagram::DiagramDefinition::parse_quick_style(
            &quick_style_xml(&limits).unwrap(),
            &limits,
        )
        .unwrap();
        crate::diagram::DiagramDefinition::parse_colors(&colors_xml(&limits).unwrap(), &limits)
            .unwrap();
        let drawing = spec.drawing_xml(5_486_400, 3_200_400, &limits).unwrap();
        assert!(drawing.contains("Ship &lt;now&gt;"));
        // Deterministic.
        assert_eq!(
            drawing,
            spec.drawing_xml(5_486_400, 3_200_400, &limits).unwrap()
        );
        assert_eq!(data, spec.data_xml(Some("rId7"), &limits).unwrap());
        // Recognized as authored, and only then.
        assert_eq!(
            Spec::from_parts(data.as_bytes(), layout_xml.as_bytes(), &limits).unwrap(),
            Some(spec.clone())
        );
        let edited = edit::set_text(data.as_bytes(), "4", "Other", &limits).unwrap();
        assert_eq!(
            Outline::read(&edited, &limits).unwrap().entries[3].text,
            "Other"
        );
        let foreign = data.replace("<dgm:spPr/>", "<dgm:spPr></dgm:spPr>");
        assert_eq!(
            Spec::from_parts(foreign.as_bytes(), layout_xml.as_bytes(), &limits).unwrap(),
            None
        );
    }
}

#[test]
fn layout_tokens_and_ids() {
    for layout in Layout::ALL {
        assert_eq!(Layout::parse(layout.name()), Some(layout));
        assert_eq!(Layout::from_unique_id(&layout.unique_id()), Some(layout));
    }
    assert_eq!(Layout::parse("PROCESS"), Some(Layout::Process));
    assert_eq!(Layout::parse("chevron"), None);
}

#[test]
fn bounds_are_enforced() {
    let mut spec = Spec::new(Layout::List);
    for index in 0..=MAX_NODES {
        spec = spec.with_node(Node::new(index.to_string()));
    }
    assert!(matches!(spec.validate(), Err(Error::Limit { .. })));
    let bad = Spec::new(Layout::List).with_node(Node::new("a\u{1}"));
    assert!(bad.validate().is_err());
    assert!(
        Spec::new(Layout::List)
            .drawing_xml(0, 10, &Limits::default())
            .is_err()
    );
    let defaults = Limits::default();
    let tiny_output = Limits::new(
        defaults.input_bytes_value(),
        1,
        defaults.xml_nodes_value(),
        defaults.xml_depth_value(),
        defaults.points_value(),
        defaults.connections_value(),
        defaults.data_text_bytes_value(),
        defaults.nodes_value(),
        defaults.node_depth_value(),
        defaults.node_text_bytes_value(),
    );
    assert!(matches!(
        Spec::new(Layout::List).data_xml(None, &tiny_output),
        Err(Error::Limit { .. })
    ));

    let tiny_input = Limits::new(
        1,
        defaults.output_bytes_value(),
        defaults.xml_nodes_value(),
        defaults.xml_depth_value(),
        defaults.points_value(),
        defaults.connections_value(),
        defaults.data_text_bytes_value(),
        defaults.nodes_value(),
        defaults.node_depth_value(),
        defaults.node_text_bytes_value(),
    );
    assert!(matches!(
        Spec::parse_outline(Layout::List, "Plan", &tiny_input),
        Err(Error::Limit {
            kind: crate::LimitKind::InputBytes,
            ..
        })
    ));
    assert!(matches!(
        Spec::new(Layout::List)
            .with_node(Node::new("Plan"))
            .outline(&tiny_output),
        Err(Error::Limit {
            kind: crate::LimitKind::OutputBytes,
            ..
        })
    ));
}

#[test]
fn empty_spec_writes_valid_parts() {
    let limits = Limits::default();
    let spec = Spec::new(Layout::Cycle);
    let data = spec.data_xml(None, &limits).unwrap();
    assert!(
        Outline::read(data.as_bytes(), &limits)
            .unwrap()
            .entries
            .is_empty()
    );
    assert!(
        spec.drawing_xml(100, 100, &limits)
            .unwrap()
            .contains("</dsp:spTree>")
    );
}

#[test]
fn set_text_keeps_foreign_formatting_and_updates_the_drawing() {
    let limits = Limits::default();
    let data = br#"<?xml version="1.0"?><dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{A}" type="doc"><dgm:prSet loTypeId="urn:x/process3"/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ru-RU"/></a:p></dgm:t></dgm:pt><dgm:pt modelId="{B}"><dgm:prSet/><dgm:spPr/><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l"/><a:r><a:rPr lang="ru-RU" b="1"/><a:t>old</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="{C}"/><dgm:pt modelId="{P}" type="pres"><dgm:prSet presAssocID="{B}" presName="node"/><dgm:spPr/></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="{X}" srcId="{A}" destId="{B}" srcOrd="0" destOrd="0"/><dgm:cxn modelId="{Y}" srcId="{A}" destId="{C}" srcOrd="1" destOrd="0"/><dgm:cxn modelId="{Z}" type="presOf" srcId="{B}" destId="{P}" srcOrd="0" destOrd="0"/></dgm:cxnLst><dgm:bg/><dgm:whole/><dgm:extLst><a:ext uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rId6" minVer="http://schemas.openxmlformats.org/drawingml/2006/diagram"/></a:ext></dgm:extLst></dgm:dataModel>"#;
    let outline = Outline::read(data, &limits).unwrap();
    assert_eq!(outline.drawing_relationship.as_deref(), Some("rId6"));
    assert_eq!(outline.entries.len(), 2);
    let edited = edit::set_text(data, "{B}", "new\nline", &limits).unwrap();
    let text = String::from_utf8(edited.clone()).unwrap();
    assert!(text.contains(r#"<a:pPr algn="l"/><a:r><a:rPr lang="ru-RU" b="1"/><a:t>new</a:t>"#));
    assert_eq!(
        Outline::read(&edited, &limits).unwrap().entries[0].text,
        "new\nline"
    );
    // An empty point gains a text body.
    let edited = edit::set_text(data, "{C}", "c", &limits).unwrap();
    assert_eq!(
        Outline::read(&edited, &limits).unwrap().entries[1].text,
        "c"
    );
    // Unknown node.
    assert!(edit::set_text(data, "{Q}", "x", &limits).is_err());
    // No-op.
    assert_eq!(
        edit::set_text(data, "{B}", "old", &limits).unwrap(),
        data.to_vec()
    );

    let drawing = br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp modelId="{P}"><dsp:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>old</a:t></a:r></a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>"#;
    let patched = edit::drawing_text(drawing, data, "{B}", "fresh", &limits)
        .unwrap()
        .unwrap();
    assert!(
        String::from_utf8(patched)
            .unwrap()
            .contains(r#"<a:rPr sz="1200"/><a:t>fresh</a:t>"#)
    );
    // A node no shape draws.
    assert_eq!(
        edit::drawing_text(drawing, data, "{C}", "x", &limits).unwrap(),
        None
    );
    // A drawing that disagrees with the data model.
    let stale = String::from_utf8(drawing.to_vec())
        .unwrap()
        .replace(">old<", ">other<");
    assert_eq!(
        edit::drawing_text(stale.as_bytes(), data, "{B}", "x", &limits).unwrap(),
        None
    );

    // Drawing references.
    let removed = edit::set_drawing_reference(data, None, &limits).unwrap();
    let removed_text = String::from_utf8(removed.clone()).unwrap();
    assert!(!removed_text.contains("extLst"));
    assert_eq!(
        Outline::read(&removed, &limits)
            .unwrap()
            .drawing_relationship,
        None
    );
    let added = edit::set_drawing_reference(&removed, Some("rId9"), &limits).unwrap();
    assert_eq!(
        Outline::read(&added, &limits)
            .unwrap()
            .drawing_relationship
            .as_deref(),
        Some("rId9")
    );
    let renamed = edit::set_drawing_reference(data, Some("rId2"), &limits).unwrap();
    assert_eq!(
        Outline::read(&renamed, &limits)
            .unwrap()
            .drawing_relationship
            .as_deref(),
        Some("rId2")
    );
    assert_eq!(
        edit::set_drawing_reference(data, Some("rId6"), &limits).unwrap(),
        data.to_vec()
    );
}

#[test]
fn outline_reads_real_files() {
    let limits = Limits::default();
    // A data model with GUID identifiers and nested nodes.
    let data = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="0" type="doc"/><dgm:pt modelId="2"><dgm:t><a:p><a:r><a:t>b</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="1"><dgm:t><a:p><a:r><a:t>a</a:t></a:r><a:br/><a:r><a:t>x</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="3"><dgm:t><a:p><a:r><a:t>c&amp;</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="9" srcId="0" destId="2" srcOrd="1"/><dgm:cxn modelId="8" srcId="0" destId="1" srcOrd="0"/><dgm:cxn modelId="7" srcId="1" destId="3" srcOrd="0"/></dgm:cxnLst></dgm:dataModel>"#;
    let outline = Outline::read(data, &limits).unwrap();
    let got: Vec<_> = outline
        .entries
        .iter()
        .map(|e| (e.text.as_str(), e.depth))
        .collect();
    assert_eq!(got, [("a\nx", 0), ("c&", 1), ("b", 0)]);
    assert!(Outline::read(b"<a/>", &limits).is_err());
}
