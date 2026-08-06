use super::{Child, ChildKind, Master};

#[test]
fn semantic_fragment_round_trip_preserves_shared_shape_children() {
    let mut master = Master::new("physical").unwrap();
    master
        .push_child(Child::new(ChildKind::Shape, "<draw:rect/>"))
        .unwrap();

    let xml = master.to_xml_fragment().unwrap();
    let reparsed = Master::from_xml_fragment(&xml).unwrap();

    assert_eq!(reparsed, master);
    assert_eq!(reparsed.children[0].kind, ChildKind::Shape);
}
