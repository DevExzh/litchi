//! Focused tests for the typed diagram data model and XML codec.
use super::*;

const TRANSITIONAL: &str = concat!(
    "<?xml version=\"1.0\"?>",
    "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" ",
    "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
    "<dgm:ptLst>",
    "<dgm:pt modelId=\"0\" type=\"doc\"><dgm:prSet loTypeId=\"urn:test/layout/process1\" ",
    "qsTypeId=\"urn:test/quickstyle/simple1\" csTypeId=\"urn:test/colors/accent1_1\"/>",
    "<dgm:spPr/><dgm:t><a:p><a:endParaRPr/></a:p></dgm:t></dgm:pt>",
    "<dgm:pt modelId=\"1\"><dgm:prSet/><dgm:t><a:p><a:r><a:t>Alpha &amp; </a:t></a:r>",
    "<a:r><a:t>Beta</a:t></a:r></a:p></dgm:t></dgm:pt>",
    "<dgm:pt modelId=\"2\"><dgm:prSet/><dgm:t><a:p><a:r><a:t>Gamma</a:t></a:r></a:p></dgm:t></dgm:pt>",
    "<dgm:pt modelId=\"3\" type=\"node\"><dgm:t><a:p><a:r><a:t>Child</a:t></a:r></a:p></dgm:t></dgm:pt>",
    "<dgm:pt modelId=\"2000\" type=\"parTrans\" cxnId=\"100\"/>",
    "<dgm:pt modelId=\"1000\" type=\"pres\"/>",
    "</dgm:ptLst>",
    "<dgm:cxnLst>",
    "<dgm:cxn modelId=\"100\" srcId=\"0\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
    "<dgm:cxn modelId=\"101\" srcId=\"0\" destId=\"2\" srcOrd=\"1\" destOrd=\"0\"/>",
    "<dgm:cxn modelId=\"102\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/>",
    "<dgm:cxn modelId=\"300\" type=\"presOf\" srcId=\"0\" destId=\"1000\" srcOrd=\"0\" destOrd=\"0\"/>",
    "</dgm:cxnLst>",
    "<dgm:bg/><dgm:whole/>",
    "</dgm:dataModel>"
);

const STRICT: &str = concat!(
    "<?xml version=\"1.0\"?>",
    "<dgm:dataModel xmlns:dgm=\"http://purl.oclc.org/ooxml/drawingml/diagram\" ",
    "xmlns:a=\"http://purl.oclc.org/ooxml/drawingml/main\">",
    "<dgm:ptLst>",
    "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000001}\" type=\"doc\"><dgm:prSet loTypeId=\"urn:test/layout/cycle2\"/></dgm:pt>",
    "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000002}\"><dgm:t><a:p><a:r><a:t>a</a:t></a:r></a:p></dgm:t></dgm:pt>",
    "<dgm:pt modelId=\"{00000000-0000-0000-0000-000000000003}\" type=\"sibTrans\" cxnId=\"{00000000-0000-0000-0000-000000000004}\"/>",
    "</dgm:ptLst>",
    "<dgm:cxnLst>",
    "<dgm:cxn modelId=\"{00000000-0000-0000-0000-000000000004}\" srcId=\"{00000000-0000-0000-0000-000000000001}\" destId=\"{00000000-0000-0000-0000-000000000002}\" srcOrd=\"0\" destOrd=\"0\"/>",
    "</dgm:cxnLst>",
    "</dgm:dataModel>"
);

#[test]
fn parses_transitional_model_with_hierarchy_and_multi_run_text() {
    let model = DiagramDataModel::parse(TRANSITIONAL).unwrap();
    assert_eq!(model.points.len(), 6);
    assert_eq!(model.connections.len(), 4);
    let root = model.document_point().unwrap();
    assert_eq!(
        root.layout_type_id.as_deref(),
        Some("urn:test/layout/process1")
    );
    assert_eq!(
        root.quick_style_type_id.as_deref(),
        Some("urn:test/quickstyle/simple1")
    );
    assert_eq!(
        model.points[4].kind,
        PointType::ParentTransition(Id::number(100))
    );
    assert_eq!(model.points[5].kind, PointType::Presentation);
    assert_eq!(
        model.connections[3].kind,
        ConnectionType::Presentation(String::new())
    );

    let tree = model.node_tree();
    assert_eq!(tree.len(), 2);
    assert_eq!(tree[0].text, "Alpha & Beta");
    assert_eq!(tree[0].depth, 0);
    assert_eq!(tree[1].text, "Gamma");
    assert_eq!(tree[1].children.len(), 1);
    assert_eq!(tree[1].children[0].text, "Child");
    assert_eq!(tree[1].children[0].depth, 1);
    assert_eq!(model.text(), "Alpha & Beta\nGamma\nChild");
}

#[test]
fn parses_strict_namespace_model() {
    let model = DiagramDataModel::parse(STRICT).unwrap();
    assert_eq!(model.points.len(), 3);
    let connection_id: Id = "{00000000-0000-0000-0000-000000000004}".parse().unwrap();
    assert_eq!(
        model.points[2].kind,
        PointType::SiblingTransition(connection_id)
    );
    assert_eq!(
        model.document_point().unwrap().layout_type_id.as_deref(),
        Some("urn:test/layout/cycle2")
    );
    let tree = model.node_tree();
    assert_eq!(tree.len(), 1);
    assert_eq!(tree[0].text, "a");
}

#[test]
fn tolerates_cycles_and_dangling_connections() {
    let xml = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst><dgm:pt modelId=\"0\" type=\"doc\"/><dgm:pt modelId=\"1\"/><dgm:pt modelId=\"2\"/></dgm:ptLst>",
        "<dgm:cxnLst>",
        "<dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"11\" srcId=\"1\" destId=\"2\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"12\" srcId=\"2\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"13\" srcId=\"0\" destId=\"9\" srcOrd=\"1\" destOrd=\"0\"/>",
        "</dgm:cxnLst></dgm:dataModel>"
    );
    let model = DiagramDataModel::parse(xml).unwrap();
    assert!(model.validate().is_err());
    let tree = model.node_tree();
    assert_eq!(tree.len(), 1);
    assert_eq!(tree[0].children.len(), 1);
    assert!(tree[0].children[0].children.is_empty());
}

#[test]
fn rejects_wrong_root_and_dtd() {
    assert!(
            DiagramDataModel::parse(
                "<dgm:layoutDef xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>"
            )
            .is_err()
        );
    assert!(
            DiagramDataModel::parse(
                "<!DOCTYPE dgm:dataModel><dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"/>"
            )
            .is_err()
        );
}

#[test]
fn rejects_missing_ids() {
    assert!(
            DiagramDataModel::parse(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:ptLst><dgm:pt type=\"doc\"/></dgm:ptLst></dgm:dataModel>"
            )
            .is_err()
        );
    assert!(
            DiagramDataModel::parse(
                "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\"><dgm:ptLst/><dgm:cxnLst><dgm:cxn modelId=\"1\" srcId=\"0\"/></dgm:cxnLst></dgm:dataModel>"
            )
            .is_err()
        );
}

#[test]
fn enforces_data_model_structure_and_unqualified_schema_attributes() {
    let namespace = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
    for invalid_xml in [
        format!("<dgm:dataModel xmlns:dgm=\"{namespace}\"/>"),
        format!(
            "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:cxnLst/><dgm:ptLst/></dgm:dataModel>"
        ),
        format!(
            "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:pt modelId=\"1\"/><dgm:ptLst/></dgm:dataModel>"
        ),
        format!(
            "<dgm:dataModel xmlns:dgm=\"{namespace}\"><dgm:ptLst><dgm:pt modelId=\"1\"><dgm:prSet/><dgm:prSet/></dgm:pt></dgm:ptLst></dgm:dataModel>"
        ),
        format!(
            "<dgm:dataModel xmlns:dgm=\"{namespace}\" xmlns:x=\"urn:extension\"><dgm:ptLst><dgm:pt x:modelId=\"1\"/></dgm:ptLst></dgm:dataModel>"
        ),
    ] {
        assert!(
            DiagramDataModel::parse(&invalid_xml).is_err(),
            "accepted structurally invalid XML: {invalid_xml}"
        );
    }
}

#[test]
fn extracts_only_drawingml_text_leaf_content() {
    let xml = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\" ",
        "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">",
        "<dgm:ptLst><dgm:pt modelId=\"1\"><dgm:t>\n",
        "  <a:p><a:r><a:t>Alpha </a:t></a:r>\n",
        "  <a:r><a:t>Beta</a:t></a:r></a:p>\n",
        "</dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>"
    );
    assert_eq!(
        DiagramDataModel::parse(xml).unwrap().points[0].text,
        "Alpha Beta"
    );
}

#[test]
fn model_id_is_a_closed_zero_allocation_wire_domain() {
    assert_eq!(" -2147483648 ".parse::<Id>().unwrap(), Id::number(i32::MIN));
    let guid: Id = "{01234567-89AB-CDEF-0123-456789ABCDEF}".parse().unwrap();
    assert_eq!(guid.to_string(), "{01234567-89AB-CDEF-0123-456789ABCDEF}");
    assert_eq!(
        guid.as_guid(),
        Some([
            0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD, 0xEF, 0x01, 0x23, 0x45, 0x67, 0x89, 0xAB,
            0xCD, 0xEF,
        ])
    );
    for invalid in [
        "2147483648",
        "id-1",
        "{01234567-89ab-CDEF-0123-456789ABCDEF}",
        "01234567-89AB-CDEF-0123-456789ABCDEF",
        "{01234567-89AB-CDEF-0123-456789ABCDE}",
    ] {
        assert!(invalid.parse::<Id>().is_err(), "accepted {invalid}");
    }
}

#[test]
fn rejects_values_outside_fixed_point_and_connection_domains() {
    let invalid_point = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst><dgm:pt modelId=\"1\" type=\"futureNode\"/></dgm:ptLst>",
        "</dgm:dataModel>"
    );
    assert!(DiagramDataModel::parse(invalid_point).is_err());

    let invalid_connection = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst/>",
        "<dgm:cxnLst><dgm:cxn modelId=\"1\" type=\"futureRelation\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/></dgm:cxnLst>",
        "</dgm:dataModel>"
    );
    assert!(DiagramDataModel::parse(invalid_connection).is_err());

    let invalid_identifier = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst><dgm:pt modelId=\"node-one\"/></dgm:ptLst>",
        "</dgm:dataModel>"
    );
    assert!(DiagramDataModel::parse(invalid_identifier).is_err());
}

#[test]
fn accepts_all_schema_connection_types_without_string_fallback() {
    let xml = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst/>",
        "<dgm:cxnLst>",
        "<dgm:cxn modelId=\"1\" type=\"presParOf\" srcId=\"2\" destId=\"3\" srcOrd=\"0\" destOrd=\"0\"/>",
        "<dgm:cxn modelId=\"4\" type=\"unknownRelationship\" srcId=\"2\" destId=\"3\" srcOrd=\"1\" destOrd=\"0\"/>",
        "</dgm:cxnLst></dgm:dataModel>"
    );
    let model = DiagramDataModel::parse(xml).unwrap();
    assert_eq!(
        model.connections[0].kind,
        ConnectionType::PresentationParent
    );
    assert_eq!(model.connections[1].kind, ConnectionType::Unknown);
}

#[test]
fn canonical_writer_round_trips_both_conformance_classes() {
    let mut model = DiagramDataModel::new();
    for point in [
        Point::new(Id::number(0), PointType::Document),
        Point::node(Id::number(1), "A & B"),
        Point::new(Id::number(2), PointType::ParentTransition(Id::number(10))),
        Point::new(Id::number(3), PointType::SiblingTransition(Id::number(10))),
        Point::new(Id::number(4), PointType::Presentation),
    ] {
        model.add_point(point).unwrap();
    }
    model
        .add_connection(Connection::new(
            Id::number(10),
            ConnectionType::parent(Id::number(2), Id::number(3)),
            Id::number(0),
            Id::number(1),
            0,
            0,
        ))
        .unwrap();
    model
        .add_connection(Connection::new(
            Id::number(11),
            ConnectionType::Presentation("urn:test/layout".to_string()),
            Id::number(1),
            Id::number(4),
            0,
            0,
        ))
        .unwrap();

    let xml = model.to_xml(Conformance::Transitional).unwrap();
    assert!(xml.contains("parTransId=\"2\" sibTransId=\"3\""));
    assert!(xml.contains("type=\"presOf\" presId=\"urn:test/layout\""));
    assert_eq!(DiagramDataModel::parse(&xml).unwrap(), model);

    let xml = model.to_xml(Conformance::Strict).unwrap();
    assert!(xml.contains(DGM_NAMESPACE_STRICT));
    assert!(xml.contains("http://purl.oclc.org/ooxml/drawingml/main"));
    assert_eq!(DiagramDataModel::parse(&xml).unwrap(), model);
}

#[test]
fn semantic_crud_guards_duplicates_and_cascades_dependencies() {
    let mut model = DiagramDataModel::new();
    model
        .add_point(Point::new(Id::number(0), PointType::Document))
        .unwrap();
    model.add_point(Point::node(Id::number(1), "one")).unwrap();
    model
        .add_point(Point::new(
            Id::number(2),
            PointType::ParentTransition(Id::number(10)),
        ))
        .unwrap();
    model
        .add_point(Point::new(
            Id::number(3),
            PointType::SiblingTransition(Id::number(10)),
        ))
        .unwrap();
    model
        .add_connection(Connection::new(
            Id::number(10),
            ConnectionType::parent(Id::number(2), Id::number(3)),
            Id::number(0),
            Id::number(1),
            0,
            0,
        ))
        .unwrap();
    assert!(model.validate().is_ok());
    assert!(
        model
            .add_point(Point::node(Id::number(10), "collision"))
            .is_err()
    );

    let mut broken_transition = model.clone();
    broken_transition.point_mut(Id::number(2)).unwrap().kind =
        PointType::ParentTransition(Id::number(99));
    assert!(broken_transition.validate().is_err());
    assert!(broken_transition.to_xml(Conformance::Transitional).is_err());

    let mut cross_domain_collision = model.clone();
    cross_domain_collision.connections[0].id = Id::number(1);
    assert!(cross_domain_collision.validate().is_err());

    let conflicting_parent = Connection::new(
        Id::number(11),
        ConnectionType::PresentationParent,
        Id::number(0),
        Id::number(1),
        1,
        0,
    );
    assert!(model.add_connection(conflicting_parent.clone()).is_err());
    model.connections.push(conflicting_parent);
    assert!(model.validate().is_err());
    model.connections.pop();

    model
        .add_point(Point::new(Id::number(4), PointType::Presentation))
        .unwrap();
    model
        .add_connection(Connection::new(
            Id::number(11),
            ConnectionType::Presentation("urn:layout/a".to_string()),
            Id::number(0),
            Id::number(4),
            0,
            0,
        ))
        .unwrap();
    assert!(
        model
            .add_connection(Connection::new(
                Id::number(12),
                ConnectionType::Presentation("urn:layout/b".to_string()),
                Id::number(1),
                Id::number(4),
                0,
                0,
            ))
            .is_err()
    );
    assert!(
        model
            .add_point(Point::node(Id::number(1), "duplicate"))
            .is_err()
    );
    model.point_mut(Id::number(1)).unwrap().text = "updated".to_string();
    assert_eq!(model.point(Id::number(1)).unwrap().text, "updated");

    let removed = model.remove_connection(Id::number(10)).unwrap();
    assert_eq!(removed.destination, Id::number(1));
    assert!(model.point(Id::number(2)).is_none());
    assert!(model.point(Id::number(3)).is_none());
}

#[test]
fn canonical_writer_enforces_xml_and_aggregate_size_budgets() {
    let mut invalid_text = DiagramDataModel::new();
    invalid_text
        .add_point(Point::node(Id::number(0), "invalid\0text"))
        .unwrap();
    assert!(invalid_text.to_xml(Conformance::Transitional).is_err());
    let mut destination = "unchanged".to_string();
    assert!(
        invalid_text
            .write_xml(&mut destination, Conformance::Transitional)
            .is_err()
    );
    assert_eq!(destination, "unchanged");

    let mut oversized = DiagramDataModel::new();
    oversized
        .add_point(Point::new(Id::number(0), PointType::Document))
        .unwrap();
    oversized
        .add_point(Point::node(Id::number(1), "child"))
        .unwrap();
    oversized
        .add_connection(Connection::new(
            Id::number(2),
            ConnectionType::Presentation("x".repeat(MAX_DATA_MODEL_XML)),
            Id::number(0),
            Id::number(1),
            0,
            0,
        ))
        .unwrap();
    assert!(oversized.validate().is_err());
    assert!(oversized.to_xml(Conformance::Transitional).is_err());
}
