//! SmartArt example - demonstrates reading and creating SmartArt diagrams.

use litchi_pptx::shape::diagram::{Builder, Graphic, Kind, Node, data_xml};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== SmartArt Example ===\n");

    // Create a List diagram
    let list = Builder::new(Kind::List)
        .layout_name("Basic Block List")
        .add_items(vec![
            "First Item",
            "Second Item",
            "Third Item",
            "Fourth Item",
        ])
        .build();

    println!("List Diagram:");
    println!("  Type: {:?}", list.diagram_type);
    println!("  Layout: {:?}", list.layout_name);
    println!("  Nodes: {}", list.node_count());
    println!("  Text:\n{}\n", list.text());

    // The typed model can round-trip through the shared DrawingML data codec.
    let list_data = data_xml(&list);
    let list_nodes = Graphic::parse_data_xml(&list_data)?;
    assert_eq!(list_nodes.len(), list.node_count());
    println!("  Data XML round-trip: {} nodes", list_nodes.len());

    // Create a Process diagram
    let process = Builder::new(Kind::Process)
        .layout_name("Basic Process")
        .add_items(vec!["Start", "Process", "Review", "Complete"])
        .build();

    println!("Process Diagram:");
    println!("  Type: {:?}", process.diagram_type);
    println!("  Nodes: {}", process.node_count());

    // Create a Hierarchy diagram (org chart)
    let mut hierarchy = Graphic::new(Kind::Hierarchy);
    hierarchy.layout_name = Some("Organization Chart".to_string());

    let mut ceo = Node::new("CEO");
    ceo.depth = 0;

    let mut vp1 = Node::new("VP Engineering");
    vp1.depth = 1;

    let mut vp2 = Node::new("VP Sales");
    vp2.depth = 1;

    ceo.add_child(vp1);
    ceo.add_child(vp2);
    hierarchy.add_node(ceo);

    println!("\nHierarchy Diagram:");
    println!("  Type: {:?}", hierarchy.diagram_type);
    println!("  Text:\n{}", hierarchy.text());

    // Test diagram type detection
    println!("\n--- Diagram Type Detection ---");
    let types = [
        ("urn:microsoft.com/office/list", Kind::List),
        ("urn:microsoft.com/office/process", Kind::Process),
        ("urn:microsoft.com/office/orgChart", Kind::Hierarchy),
        ("urn:microsoft.com/office/venn", Kind::Relationship),
        ("urn:microsoft.com/office/matrix", Kind::Matrix),
        ("urn:microsoft.com/office/pyramid", Kind::Pyramid),
    ];

    for (uri, expected) in types {
        let detected = Kind::from_layout_uri(uri);
        println!("  {} -> {:?} (expected: {:?})", uri, detected, expected);
        assert_eq!(detected, expected);
    }

    println!("\n✅ SmartArt example completed successfully!");
    Ok(())
}
