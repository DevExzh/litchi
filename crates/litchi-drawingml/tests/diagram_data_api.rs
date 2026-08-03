use litchi_drawingml::diagram::data::{
    Conformance, Connection, ConnectionType, DiagramDataModel, Id, Point, PointType,
};

#[test]
fn typed_diagram_data_crud_round_trips_without_raw_identifier_strings() {
    let root = Id::number(0);
    let child = Id::number(1);
    let connection = Id::number(10);
    let parent_transition = Id::number(20);
    let sibling_transition = Id::number(21);

    let mut model = DiagramDataModel::new();
    model
        .add_point(Point::new(root, PointType::Document))
        .unwrap();
    model.add_point(Point::node(child, "A & B")).unwrap();
    model
        .add_point(Point::new(
            parent_transition,
            PointType::ParentTransition(connection),
        ))
        .unwrap();
    model
        .add_point(Point::new(
            sibling_transition,
            PointType::SiblingTransition(connection),
        ))
        .unwrap();
    model
        .add_connection(Connection::new(
            connection,
            ConnectionType::parent(parent_transition, sibling_transition),
            root,
            child,
            0,
            0,
        ))
        .unwrap();

    assert_eq!(model.point(child).unwrap().text, "A & B");
    let xml = model.to_xml(Conformance::Strict).unwrap();
    assert!(xml.contains("parTransId=\"20\" sibTransId=\"21\""));
    assert!(xml.contains("A &amp; B"));
    assert_eq!(DiagramDataModel::parse(&xml).unwrap(), model);
}

#[test]
fn invalid_fixed_domain_values_never_reach_the_public_model() {
    assert!("{not-a-guid}".parse::<Id>().is_err());
    let xml = concat!(
        "<dgm:dataModel xmlns:dgm=\"http://schemas.openxmlformats.org/drawingml/2006/diagram\">",
        "<dgm:ptLst><dgm:pt modelId=\"1\" type=\"custom\"/></dgm:ptLst>",
        "</dgm:dataModel>"
    );
    assert!(DiagramDataModel::parse(xml).is_err());
}
