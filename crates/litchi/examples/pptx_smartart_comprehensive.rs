//! Comprehensive SmartArt Example
//!
//! Demonstrates creating various SmartArt/diagram types and layouts.
//! SmartArt graphics provide visual representations of information and ideas.

use litchi_pptx::Package;
use litchi_pptx::shape;
use litchi_pptx::shape::diagram::{
    Builder, Graphic, Kind, Node, colors_xml, data_xml, drawing_xml, graphic_frame, layout_xml,
    quickstyle_xml,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive SmartArt Example ===\n");

    // Diagram models and their five DrawingML part codecs are package-independent.
    // The PPTX package is used below for the canonical package boundary check;
    // no retired mutable presentation writer is needed for SmartArt authoring.
    println!("Creating typed SmartArt examples...\n");

    test_list_diagrams()?;
    test_process_diagrams()?;
    test_cycle_diagrams()?;
    test_hierarchy_diagrams()?;
    test_relationship_diagrams()?;
    test_matrix_diagrams()?;
    test_pyramid_diagrams()?;
    test_complex_hierarchies()?;

    let mut pkg = Package::new()?;
    let bytes = pkg.to_bytes()?;
    std::fs::write("smartart_comprehensive.pptx", &bytes)?;
    let reopened = Package::from_bytes(&bytes)?;
    let slide_count = reopened.presentation()?.slide_count()?;
    let part_count = reopened.with_opc(|opc| Ok(opc.iter_parts().count()))?;
    assert_eq!(slide_count, 0);
    assert!(part_count > 0);
    validate_graphic_frame()?;
    println!("\n✓ Saved: smartart_comprehensive.pptx");

    println!("\n=== All SmartArt examples complete! ===");
    println!(
        "\nExercised {} typed diagram cases across the canonical codecs.",
        8
    );
    println!("\nSmartArt XML components have been generated:");
    println!("  ✓ List Diagrams (Block lists, vertical lists)");
    println!("  ✓ Process Diagrams (Basic process, continuous block)");
    println!("  ✓ Cycle Diagrams (PDCA, circular cycles)");
    println!("  ✓ Hierarchy Diagrams (Org charts, tree structures)");
    println!("  ✓ Relationship Diagrams (Venn, radial)");
    println!("  ✓ Matrix Diagrams (2x2, grid)");
    println!("  ✓ Pyramid Diagrams (Hierarchical levels)");
    println!("  ✓ Complex Hierarchies (Multi-level organizations)");
    println!("\nNote: SmartArt XML generation is demonstrated.");
    println!("Full SmartArt integration requires additional diagram part infrastructure.");

    Ok(())
}

/// Test 1: List Diagrams
/// Show non-sequential or grouped information
fn test_list_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 1: List Diagrams");
    println!("---------------------");

    // Basic bullet list
    let bullet_list = Builder::new(Kind::List)
        .layout_name("Basic Block List")
        .add_items(vec![
            "Define project scope",
            "Allocate resources",
            "Set milestones",
            "Monitor progress",
            "Deliver results",
        ])
        .build();

    let xml = data_xml(&bullet_list);
    println!(
        "  ✓ Basic Block List: {} nodes, {} bytes XML",
        bullet_list.node_count(),
        xml.len()
    );

    // Feature list with descriptions
    let feature_list = Builder::new(Kind::List)
        .layout_name("Vertical Box List")
        .add_items(vec![
            "Real-time Analytics: Monitor performance metrics instantly",
            "Automated Workflows: Streamline repetitive tasks",
            "Cloud Integration: Seamless data synchronization",
            "Mobile Access: Work from anywhere",
            "Security & Compliance: Enterprise-grade protection",
        ])
        .build();

    let xml = data_xml(&feature_list);
    println!(
        "  ✓ Feature List: {} items, {} bytes XML",
        feature_list.node_count(),
        xml.len()
    );

    // Grouped list (categories)
    let categories = Builder::new(Kind::List)
        .layout_name("Grouped List")
        .add_items(vec![
            "Frontend: React, Vue, Angular",
            "Backend: Node.js, Python, Java",
            "Database: PostgreSQL, MongoDB, Redis",
            "DevOps: Docker, Kubernetes, CI/CD",
        ])
        .build();

    let xml = data_xml(&categories);
    println!(
        "  ✓ Technology Stack: {} categories, {} bytes XML\n",
        categories.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 2: Process Diagrams
/// Show steps in a process or timeline
fn test_process_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 2: Process Diagrams");
    println!("------------------------");

    // Basic process flow
    let basic_process = Builder::new(Kind::Process)
        .layout_name("Basic Process")
        .add_items(vec![
            "Research", "Plan", "Design", "Develop", "Test", "Deploy",
        ])
        .build();

    let xml = data_xml(&basic_process);
    println!(
        "  ✓ Software Development Process: {} steps, {} bytes XML",
        basic_process.node_count(),
        xml.len()
    );

    // Customer journey
    let customer_journey = Builder::new(Kind::Process)
        .layout_name("Continuous Block Process")
        .add_items(vec![
            "Awareness",
            "Consideration",
            "Decision",
            "Purchase",
            "Retention",
            "Advocacy",
        ])
        .build();

    let xml = data_xml(&customer_journey);
    println!(
        "  ✓ Customer Journey: {} stages, {} bytes XML",
        customer_journey.node_count(),
        xml.len()
    );

    // Manufacturing pipeline
    let pipeline = Builder::new(Kind::Process)
        .layout_name("Step-Up Process")
        .add_items(vec![
            "Raw Materials",
            "Processing",
            "Assembly",
            "Quality Control",
            "Packaging",
            "Distribution",
            "Delivery",
        ])
        .build();

    let xml = data_xml(&pipeline);
    println!(
        "  ✓ Manufacturing Pipeline: {} steps, {} bytes XML\n",
        pipeline.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 3: Cycle Diagrams
/// Show continuous or repeating processes
fn test_cycle_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 3: Cycle Diagrams");
    println!("----------------------");

    // Continuous improvement cycle
    let pdca_cycle = Builder::new(Kind::Cycle)
        .layout_name("Basic Cycle")
        .add_items(vec![
            "Plan: Identify opportunities",
            "Do: Implement changes",
            "Check: Evaluate results",
            "Act: Standardize or adjust",
        ])
        .build();

    let xml = data_xml(&pdca_cycle);
    println!(
        "  ✓ PDCA Cycle: {} phases, {} bytes XML",
        pdca_cycle.node_count(),
        xml.len()
    );

    // Agile sprint cycle
    let agile_sprint = Builder::new(Kind::Cycle)
        .layout_name("Circular Cycle")
        .add_items(vec![
            "Sprint Planning",
            "Daily Standups",
            "Development",
            "Testing & Review",
            "Sprint Retrospective",
        ])
        .build();

    let xml = data_xml(&agile_sprint);
    println!(
        "  ✓ Agile Sprint: {} activities, {} bytes XML",
        agile_sprint.node_count(),
        xml.len()
    );

    // Product lifecycle
    let lifecycle = Builder::new(Kind::Cycle)
        .layout_name("Block Cycle")
        .add_items(vec![
            "Introduction",
            "Growth",
            "Maturity",
            "Decline",
            "Renewal/Retirement",
        ])
        .build();

    let xml = data_xml(&lifecycle);
    println!(
        "  ✓ Product Lifecycle: {} stages, {} bytes XML\n",
        lifecycle.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 4: Hierarchy Diagrams
/// Show organizational structure or rankings
fn test_hierarchy_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 4: Hierarchy Diagrams");
    println!("--------------------------");

    // Simple org chart
    let org_chart = Builder::new(Kind::Hierarchy)
        .layout_name("Organization Chart")
        .add_items(vec!["CEO", "CFO", "CTO", "COO"])
        .build();

    let xml = data_xml(&org_chart);
    println!(
        "  ✓ Executive Team: {} positions, {} bytes XML",
        org_chart.node_count(),
        xml.len()
    );

    // Detailed hierarchy with builder
    let mut company_org = Graphic::new(Kind::Hierarchy);
    company_org.layout_name = Some("Hierarchy".to_string());

    let mut ceo = Node::new("CEO - Jane Smith");
    ceo.depth = 0;

    let mut cfo = Node::new("CFO - John Doe");
    cfo.depth = 1;
    cfo.add_child(Node::new("Finance Manager"));
    cfo.add_child(Node::new("Accounting Manager"));

    let mut cto = Node::new("CTO - Alice Johnson");
    cto.depth = 1;
    cto.add_child(Node::new("Engineering Manager"));
    cto.add_child(Node::new("DevOps Manager"));
    cto.add_child(Node::new("QA Manager"));

    let mut coo = Node::new("COO - Bob Wilson");
    coo.depth = 1;
    coo.add_child(Node::new("Operations Manager"));
    coo.add_child(Node::new("Supply Chain Manager"));

    ceo.add_child(cfo);
    ceo.add_child(cto);
    ceo.add_child(coo);
    company_org.add_node(ceo);

    let xml = data_xml(&company_org);
    println!(
        "  ✓ Company Organization: {} total nodes, {} bytes XML",
        company_org.node_count(),
        xml.len()
    );

    // Technology stack hierarchy
    let tech_stack = Builder::new(Kind::Hierarchy)
        .layout_name("Hierarchy List")
        .add_items(vec![
            "Application Layer",
            "Business Logic Layer",
            "Data Access Layer",
            "Database Layer",
            "Infrastructure Layer",
        ])
        .build();

    let xml = data_xml(&tech_stack);
    println!(
        "  ✓ Tech Stack Layers: {} layers, {} bytes XML\n",
        tech_stack.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 5: Relationship Diagrams
/// Show connections and relationships
fn test_relationship_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 5: Relationship Diagrams");
    println!("-----------------------------");

    // Venn diagram (overlapping concepts)
    let venn = Builder::new(Kind::Relationship)
        .layout_name("Basic Venn")
        .add_items(vec!["Skills", "Passion", "Market Need"])
        .build();

    let xml = data_xml(&venn);
    println!(
        "  ✓ Finding Your Niche: {} circles, {} bytes XML",
        venn.node_count(),
        xml.len()
    );

    // Balanced relationship
    let balance = Builder::new(Kind::Relationship)
        .layout_name("Converging Radial")
        .add_items(vec!["Quality", "Speed", "Cost", "Scope"])
        .build();

    let xml = data_xml(&balance);
    println!(
        "  ✓ Project Balance: {} factors, {} bytes XML",
        balance.node_count(),
        xml.len()
    );

    // Interconnected systems
    let systems = Builder::new(Kind::Relationship)
        .layout_name("Basic Radial")
        .add_items(vec![
            "Core Platform",
            "API Gateway",
            "Microservices",
            "Data Layer",
            "Client Apps",
        ])
        .build();

    let xml = data_xml(&systems);
    println!(
        "  ✓ System Architecture: {} components, {} bytes XML\n",
        systems.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 6: Matrix Diagrams
/// Show how parts relate to a whole
fn test_matrix_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 6: Matrix Diagrams");
    println!("-----------------------");

    // 2x2 matrix (priority matrix)
    let priority_matrix = Builder::new(Kind::Matrix)
        .layout_name("Basic Matrix")
        .add_items(vec![
            "Urgent & Important",
            "Not Urgent & Important",
            "Urgent & Not Important",
            "Not Urgent & Not Important",
        ])
        .build();

    let xml = data_xml(&priority_matrix);
    println!(
        "  ✓ Priority Matrix: {} quadrants, {} bytes XML",
        priority_matrix.node_count(),
        xml.len()
    );

    // Product-market fit matrix
    let market_fit = Builder::new(Kind::Matrix)
        .layout_name("Grid Matrix")
        .add_items(vec![
            "High Value, Easy to Build",
            "High Value, Hard to Build",
            "Low Value, Easy to Build",
            "Low Value, Hard to Build",
        ])
        .build();

    let xml = data_xml(&market_fit);
    println!(
        "  ✓ Feature Prioritization: {} categories, {} bytes XML\n",
        market_fit.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 7: Pyramid Diagrams
/// Show proportional or hierarchical relationships
fn test_pyramid_diagrams() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 7: Pyramid Diagrams");
    println!("------------------------");

    // Needs hierarchy
    let needs_pyramid = Builder::new(Kind::Pyramid)
        .layout_name("Basic Pyramid")
        .add_items(vec![
            "Self-Actualization",
            "Esteem",
            "Love & Belonging",
            "Safety",
            "Physiological Needs",
        ])
        .build();

    let xml = data_xml(&needs_pyramid);
    println!(
        "  ✓ Maslow's Hierarchy: {} levels, {} bytes XML",
        needs_pyramid.node_count(),
        xml.len()
    );

    // Business pyramid
    let business_pyramid = Builder::new(Kind::Pyramid)
        .layout_name("Inverted Pyramid")
        .add_items(vec![
            "Vision & Strategy",
            "Goals & Objectives",
            "Tactics & Plans",
            "Daily Activities",
        ])
        .build();

    let xml = data_xml(&business_pyramid);
    println!(
        "  ✓ Strategic Planning: {} tiers, {} bytes XML",
        business_pyramid.node_count(),
        xml.len()
    );

    // Knowledge pyramid
    let knowledge = Builder::new(Kind::Pyramid)
        .layout_name("Segmented Pyramid")
        .add_items(vec!["Wisdom", "Knowledge", "Information", "Data"])
        .build();

    let xml = data_xml(&knowledge);
    println!(
        "  ✓ Knowledge Pyramid: {} levels, {} bytes XML\n",
        knowledge.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 8: Complex Hierarchies
/// Advanced organizational structures
fn test_complex_hierarchies() -> Result<(), Box<dyn std::error::Error>> {
    println!("Test 8: Complex Hierarchies");
    println!("---------------------------");

    // Detailed multi-level org chart
    let mut enterprise_org = Graphic::new(Kind::Hierarchy);
    enterprise_org.layout_name = Some("Organization Chart".to_string());

    let mut board = Node::new("Board of Directors");
    board.depth = 0;

    let mut ceo = Node::new("Chief Executive Officer");
    ceo.depth = 1;

    // Executive team with departments
    let mut cfo = Node::new("Chief Financial Officer");
    cfo.depth = 2;
    cfo.add_child(Node::new("VP Finance"));
    cfo.add_child(Node::new("VP Accounting"));
    cfo.add_child(Node::new("VP Treasury"));

    let mut cto = Node::new("Chief Technology Officer");
    cto.depth = 2;
    let mut vp_eng = Node::new("VP Engineering");
    vp_eng.add_child(Node::new("Director of Backend"));
    vp_eng.add_child(Node::new("Director of Frontend"));
    vp_eng.add_child(Node::new("Director of Mobile"));
    cto.add_child(vp_eng);
    cto.add_child(Node::new("VP Infrastructure"));
    cto.add_child(Node::new("VP Security"));

    let mut cmo = Node::new("Chief Marketing Officer");
    cmo.depth = 2;
    cmo.add_child(Node::new("VP Product Marketing"));
    cmo.add_child(Node::new("VP Content Marketing"));
    cmo.add_child(Node::new("VP Demand Generation"));

    let mut coo = Node::new("Chief Operating Officer");
    coo.depth = 2;
    coo.add_child(Node::new("VP Operations"));
    coo.add_child(Node::new("VP Customer Success"));
    coo.add_child(Node::new("VP Support"));

    ceo.add_child(cfo);
    ceo.add_child(cto);
    ceo.add_child(cmo);
    ceo.add_child(coo);
    board.add_child(ceo);
    enterprise_org.add_node(board);

    let data_xml = data_xml(&enterprise_org);
    let layout_xml = layout_xml(&enterprise_org);
    let colors_xml = colors_xml();
    let style_xml = quickstyle_xml();
    let drawing_xml = drawing_xml(&enterprise_org, 0, 0, 7_315_200, 4_000_000);

    println!("  ✓ Enterprise Organization Chart:");
    println!("     - Data XML: {} bytes", data_xml.len());
    println!("     - Layout XML: {} bytes", layout_xml.len());
    println!("     - Colors XML: {} bytes", colors_xml.len());
    println!("     - Style XML: {} bytes", style_xml.len());
    println!("     - Drawing XML: {} bytes", drawing_xml.len());
    println!("     - Total nodes: {}", enterprise_org.node_count());

    let all_text = enterprise_org.text();
    println!("     - Contains {} characters of text", all_text.len());
    println!();

    Ok(())
}

fn validate_graphic_frame() -> Result<(), Box<dyn std::error::Error>> {
    let frame = graphic_frame(2, 914_400, 1_600_000, 7_315_200, 3_800_000, "rId10");
    let xml = format!(
        r#"<p:spTree xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">{frame}</p:spTree>"#
    );
    let scene = shape::read(xml.as_bytes())?;
    assert!(matches!(scene.at(0)?, shape::Shape::Diagram(_)));
    Ok(())
}
