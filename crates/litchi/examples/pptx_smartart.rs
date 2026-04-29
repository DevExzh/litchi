//! SmartArt example - demonstrates reading and creating SmartArt diagrams.

use litchi::ooxml::pptx::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};

fn main() {
    println!("=== SmartArt Example ===\n");

    // Create a List diagram
    let list = SmartArtBuilder::new(DiagramType::List)
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

    // Create a Process diagram
    let process = SmartArtBuilder::new(DiagramType::Process)
        .layout_name("Basic Process")
        .add_items(vec!["Start", "Process", "Review", "Complete"])
        .build();

    println!("Process Diagram:");
    println!("  Type: {:?}", process.diagram_type);
    println!("  Nodes: {}", process.node_count());

    // Create a Hierarchy diagram (org chart)
    let mut hierarchy = SmartArt::new(DiagramType::Hierarchy);
    hierarchy.layout_name = Some("Organization Chart".to_string());

    let mut ceo = DiagramNode::new("CEO");
    ceo.depth = 0;

    let mut vp1 = DiagramNode::new("VP Engineering");
    vp1.depth = 1;

    let mut vp2 = DiagramNode::new("VP Sales");
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
        ("urn:microsoft.com/office/list", DiagramType::List),
        ("urn:microsoft.com/office/process", DiagramType::Process),
        ("urn:microsoft.com/office/orgChart", DiagramType::Hierarchy),
        ("urn:microsoft.com/office/venn", DiagramType::Relationship),
        ("urn:microsoft.com/office/matrix", DiagramType::Matrix),
        ("urn:microsoft.com/office/pyramid", DiagramType::Pyramid),
    ];

    for (uri, expected) in types {
        let detected = DiagramType::from_layout_uri(uri);
        println!("  {} -> {:?} (expected: {:?})", uri, detected, expected);
        assert_eq!(detected, expected);
    }

    println!("\n✅ SmartArt example completed successfully!");
}
