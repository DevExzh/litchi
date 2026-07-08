//! Comprehensive SmartArt Example
//!
//! Demonstrates creating various SmartArt/diagram types and layouts.
//! SmartArt graphics provide visual representations of information and ideas.

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::smartart::{
    DiagramNode, DiagramType, SmartArt, SmartArtBuilder, generate_smartart_colors_xml,
    generate_smartart_data_xml, generate_smartart_layout_xml, generate_smartart_quickstyle_xml,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== Comprehensive SmartArt Example ===\n");

    // Create presentation with SmartArt documentation
    println!("Creating presentation with SmartArt examples...\n");

    let mut pkg = Package::new()?;
    {
        let pres = pkg.presentation_mut()?;

        // Test different diagram types and add slides for each
        test_list_diagrams(pres)?;
        test_process_diagrams(pres)?;
        test_cycle_diagrams(pres)?;
        test_hierarchy_diagrams(pres)?;
        test_relationship_diagrams(pres)?;
        test_matrix_diagrams(pres)?;
        test_pyramid_diagrams(pres)?;
        test_complex_hierarchies(pres)?;
    }
    pkg.save("smartart_comprehensive.pptx")?;
    println!("\n✓ Saved: smartart_comprehensive.pptx");

    println!("\n=== All SmartArt examples complete! ===");
    println!(
        "\nPresentation created with {} slides documenting diagram types.",
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
fn test_list_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 1: List Diagrams");
    println!("---------------------");

    // Basic bullet list
    let bullet_list = SmartArtBuilder::new(DiagramType::List)
        .layout_name("Basic Block List")
        .add_items(vec![
            "Define project scope",
            "Allocate resources",
            "Set milestones",
            "Monitor progress",
            "Deliver results",
        ])
        .build();

    let xml = generate_smartart_data_xml(&bullet_list);
    println!(
        "  ✓ Basic Block List: {} nodes, {} bytes XML",
        bullet_list.node_count(),
        xml.len()
    );

    // Feature list with descriptions
    let feature_list = SmartArtBuilder::new(DiagramType::List)
        .layout_name("Vertical Box List")
        .add_items(vec![
            "Real-time Analytics: Monitor performance metrics instantly",
            "Automated Workflows: Streamline repetitive tasks",
            "Cloud Integration: Seamless data synchronization",
            "Mobile Access: Work from anywhere",
            "Security & Compliance: Enterprise-grade protection",
        ])
        .build();

    let xml = generate_smartart_data_xml(&feature_list);
    println!(
        "  ✓ Feature List: {} items, {} bytes XML",
        feature_list.node_count(),
        xml.len()
    );

    // Grouped list (categories)
    let categories = SmartArtBuilder::new(DiagramType::List)
        .layout_name("Grouped List")
        .add_items(vec![
            "Frontend: React, Vue, Angular",
            "Backend: Node.js, Python, Java",
            "Database: PostgreSQL, MongoDB, Redis",
            "DevOps: Docker, Kubernetes, CI/CD",
        ])
        .build();

    let xml = generate_smartart_data_xml(&categories);
    println!(
        "  ✓ Technology Stack: {} categories, {} bytes XML\n",
        categories.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 2: Process Diagrams
/// Show steps in a process or timeline
fn test_process_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 2: Process Diagrams");
    println!("------------------------");

    // Basic process flow
    let basic_process = SmartArtBuilder::new(DiagramType::Process)
        .layout_name("Basic Process")
        .add_items(vec![
            "Research", "Plan", "Design", "Develop", "Test", "Deploy",
        ])
        .build();

    let xml = generate_smartart_data_xml(&basic_process);
    println!(
        "  ✓ Software Development Process: {} steps, {} bytes XML",
        basic_process.node_count(),
        xml.len()
    );

    // Customer journey
    let customer_journey = SmartArtBuilder::new(DiagramType::Process)
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

    let xml = generate_smartart_data_xml(&customer_journey);
    println!(
        "  ✓ Customer Journey: {} stages, {} bytes XML",
        customer_journey.node_count(),
        xml.len()
    );

    // Manufacturing pipeline
    let pipeline = SmartArtBuilder::new(DiagramType::Process)
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

    let xml = generate_smartart_data_xml(&pipeline);
    println!(
        "  ✓ Manufacturing Pipeline: {} steps, {} bytes XML\n",
        pipeline.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 3: Cycle Diagrams
/// Show continuous or repeating processes
fn test_cycle_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 3: Cycle Diagrams");
    println!("----------------------");

    // Continuous improvement cycle
    let pdca_cycle = SmartArtBuilder::new(DiagramType::Cycle)
        .layout_name("Basic Cycle")
        .add_items(vec![
            "Plan: Identify opportunities",
            "Do: Implement changes",
            "Check: Evaluate results",
            "Act: Standardize or adjust",
        ])
        .build();

    let xml = generate_smartart_data_xml(&pdca_cycle);
    println!(
        "  ✓ PDCA Cycle: {} phases, {} bytes XML",
        pdca_cycle.node_count(),
        xml.len()
    );

    // Agile sprint cycle
    let agile_sprint = SmartArtBuilder::new(DiagramType::Cycle)
        .layout_name("Circular Cycle")
        .add_items(vec![
            "Sprint Planning",
            "Daily Standups",
            "Development",
            "Testing & Review",
            "Sprint Retrospective",
        ])
        .build();

    let xml = generate_smartart_data_xml(&agile_sprint);
    println!(
        "  ✓ Agile Sprint: {} activities, {} bytes XML",
        agile_sprint.node_count(),
        xml.len()
    );

    // Product lifecycle
    let lifecycle = SmartArtBuilder::new(DiagramType::Cycle)
        .layout_name("Block Cycle")
        .add_items(vec![
            "Introduction",
            "Growth",
            "Maturity",
            "Decline",
            "Renewal/Retirement",
        ])
        .build();

    let xml = generate_smartart_data_xml(&lifecycle);
    println!(
        "  ✓ Product Lifecycle: {} stages, {} bytes XML\n",
        lifecycle.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 4: Hierarchy Diagrams
/// Show organizational structure or rankings
fn test_hierarchy_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 4: Hierarchy Diagrams");
    println!("--------------------------");

    // Simple org chart
    let org_chart = SmartArtBuilder::new(DiagramType::Hierarchy)
        .layout_name("Organization Chart")
        .add_items(vec!["CEO", "CFO", "CTO", "COO"])
        .build();

    let xml = generate_smartart_data_xml(&org_chart);
    println!(
        "  ✓ Executive Team: {} positions, {} bytes XML",
        org_chart.node_count(),
        xml.len()
    );

    // Detailed hierarchy with builder
    let mut company_org = SmartArt::new(DiagramType::Hierarchy);
    company_org.layout_name = Some("Hierarchy".to_string());

    let mut ceo = DiagramNode::new("CEO - Jane Smith");
    ceo.depth = 0;

    let mut cfo = DiagramNode::new("CFO - John Doe");
    cfo.depth = 1;
    cfo.add_child(DiagramNode::new("Finance Manager"));
    cfo.add_child(DiagramNode::new("Accounting Manager"));

    let mut cto = DiagramNode::new("CTO - Alice Johnson");
    cto.depth = 1;
    cto.add_child(DiagramNode::new("Engineering Manager"));
    cto.add_child(DiagramNode::new("DevOps Manager"));
    cto.add_child(DiagramNode::new("QA Manager"));

    let mut coo = DiagramNode::new("COO - Bob Wilson");
    coo.depth = 1;
    coo.add_child(DiagramNode::new("Operations Manager"));
    coo.add_child(DiagramNode::new("Supply Chain Manager"));

    ceo.add_child(cfo);
    ceo.add_child(cto);
    ceo.add_child(coo);
    company_org.add_node(ceo);

    let xml = generate_smartart_data_xml(&company_org);
    println!(
        "  ✓ Company Organization: {} total nodes, {} bytes XML",
        company_org.node_count(),
        xml.len()
    );

    // Technology stack hierarchy
    let tech_stack = SmartArtBuilder::new(DiagramType::Hierarchy)
        .layout_name("Hierarchy List")
        .add_items(vec![
            "Application Layer",
            "Business Logic Layer",
            "Data Access Layer",
            "Database Layer",
            "Infrastructure Layer",
        ])
        .build();

    let xml = generate_smartart_data_xml(&tech_stack);
    println!(
        "  ✓ Tech Stack Layers: {} layers, {} bytes XML\n",
        tech_stack.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 5: Relationship Diagrams
/// Show connections and relationships
fn test_relationship_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 5: Relationship Diagrams");
    println!("-----------------------------");

    // Venn diagram (overlapping concepts)
    let venn = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("Basic Venn")
        .add_items(vec!["Skills", "Passion", "Market Need"])
        .build();

    let xml = generate_smartart_data_xml(&venn);
    println!(
        "  ✓ Finding Your Niche: {} circles, {} bytes XML",
        venn.node_count(),
        xml.len()
    );

    // Balanced relationship
    let balance = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("Converging Radial")
        .add_items(vec!["Quality", "Speed", "Cost", "Scope"])
        .build();

    let xml = generate_smartart_data_xml(&balance);
    println!(
        "  ✓ Project Balance: {} factors, {} bytes XML",
        balance.node_count(),
        xml.len()
    );

    // Interconnected systems
    let systems = SmartArtBuilder::new(DiagramType::Relationship)
        .layout_name("Basic Radial")
        .add_items(vec![
            "Core Platform",
            "API Gateway",
            "Microservices",
            "Data Layer",
            "Client Apps",
        ])
        .build();

    let xml = generate_smartart_data_xml(&systems);
    println!(
        "  ✓ System Architecture: {} components, {} bytes XML\n",
        systems.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 6: Matrix Diagrams
/// Show how parts relate to a whole
fn test_matrix_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 6: Matrix Diagrams");
    println!("-----------------------");

    // 2x2 matrix (priority matrix)
    let priority_matrix = SmartArtBuilder::new(DiagramType::Matrix)
        .layout_name("Basic Matrix")
        .add_items(vec![
            "Urgent & Important",
            "Not Urgent & Important",
            "Urgent & Not Important",
            "Not Urgent & Not Important",
        ])
        .build();

    let xml = generate_smartart_data_xml(&priority_matrix);
    println!(
        "  ✓ Priority Matrix: {} quadrants, {} bytes XML",
        priority_matrix.node_count(),
        xml.len()
    );

    // Product-market fit matrix
    let market_fit = SmartArtBuilder::new(DiagramType::Matrix)
        .layout_name("Grid Matrix")
        .add_items(vec![
            "High Value, Easy to Build",
            "High Value, Hard to Build",
            "Low Value, Easy to Build",
            "Low Value, Hard to Build",
        ])
        .build();

    let xml = generate_smartart_data_xml(&market_fit);
    println!(
        "  ✓ Feature Prioritization: {} categories, {} bytes XML\n",
        market_fit.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 7: Pyramid Diagrams
/// Show proportional or hierarchical relationships
fn test_pyramid_diagrams(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 7: Pyramid Diagrams");
    println!("------------------------");

    // Needs hierarchy
    let needs_pyramid = SmartArtBuilder::new(DiagramType::Pyramid)
        .layout_name("Basic Pyramid")
        .add_items(vec![
            "Self-Actualization",
            "Esteem",
            "Love & Belonging",
            "Safety",
            "Physiological Needs",
        ])
        .build();

    let xml = generate_smartart_data_xml(&needs_pyramid);
    println!(
        "  ✓ Maslow's Hierarchy: {} levels, {} bytes XML",
        needs_pyramid.node_count(),
        xml.len()
    );

    // Business pyramid
    let business_pyramid = SmartArtBuilder::new(DiagramType::Pyramid)
        .layout_name("Inverted Pyramid")
        .add_items(vec![
            "Vision & Strategy",
            "Goals & Objectives",
            "Tactics & Plans",
            "Daily Activities",
        ])
        .build();

    let xml = generate_smartart_data_xml(&business_pyramid);
    println!(
        "  ✓ Strategic Planning: {} tiers, {} bytes XML",
        business_pyramid.node_count(),
        xml.len()
    );

    // Knowledge pyramid
    let knowledge = SmartArtBuilder::new(DiagramType::Pyramid)
        .layout_name("Segmented Pyramid")
        .add_items(vec!["Wisdom", "Knowledge", "Information", "Data"])
        .build();

    let xml = generate_smartart_data_xml(&knowledge);
    println!(
        "  ✓ Knowledge Pyramid: {} levels, {} bytes XML\n",
        knowledge.node_count(),
        xml.len()
    );

    Ok(())
}

/// Test 8: Complex Hierarchies
/// Advanced organizational structures
fn test_complex_hierarchies(
    pres: &mut litchi::ooxml::pptx::writer::pres::MutablePresentation,
) -> Result<(), Box<dyn std::error::Error>> {
    let slide = pres.add_slide()?;
    slide.set_title("SmartArt Diagram Types");
    println!("Test 8: Complex Hierarchies");
    println!("---------------------------");

    // Detailed multi-level org chart
    let mut enterprise_org = SmartArt::new(DiagramType::Hierarchy);
    enterprise_org.layout_name = Some("Organization Chart".to_string());

    let mut board = DiagramNode::new("Board of Directors");
    board.depth = 0;

    let mut ceo = DiagramNode::new("Chief Executive Officer");
    ceo.depth = 1;

    // Executive team with departments
    let mut cfo = DiagramNode::new("Chief Financial Officer");
    cfo.depth = 2;
    cfo.add_child(DiagramNode::new("VP Finance"));
    cfo.add_child(DiagramNode::new("VP Accounting"));
    cfo.add_child(DiagramNode::new("VP Treasury"));

    let mut cto = DiagramNode::new("Chief Technology Officer");
    cto.depth = 2;
    let mut vp_eng = DiagramNode::new("VP Engineering");
    vp_eng.add_child(DiagramNode::new("Director of Backend"));
    vp_eng.add_child(DiagramNode::new("Director of Frontend"));
    vp_eng.add_child(DiagramNode::new("Director of Mobile"));
    cto.add_child(vp_eng);
    cto.add_child(DiagramNode::new("VP Infrastructure"));
    cto.add_child(DiagramNode::new("VP Security"));

    let mut cmo = DiagramNode::new("Chief Marketing Officer");
    cmo.depth = 2;
    cmo.add_child(DiagramNode::new("VP Product Marketing"));
    cmo.add_child(DiagramNode::new("VP Content Marketing"));
    cmo.add_child(DiagramNode::new("VP Demand Generation"));

    let mut coo = DiagramNode::new("Chief Operating Officer");
    coo.depth = 2;
    coo.add_child(DiagramNode::new("VP Operations"));
    coo.add_child(DiagramNode::new("VP Customer Success"));
    coo.add_child(DiagramNode::new("VP Support"));

    ceo.add_child(cfo);
    ceo.add_child(cto);
    ceo.add_child(cmo);
    ceo.add_child(coo);
    board.add_child(ceo);
    enterprise_org.add_node(board);

    let data_xml = generate_smartart_data_xml(&enterprise_org);
    let layout_xml = generate_smartart_layout_xml(&enterprise_org);
    let colors_xml = generate_smartart_colors_xml();
    let style_xml = generate_smartart_quickstyle_xml();

    println!("  ✓ Enterprise Organization Chart:");
    println!("     - Data XML: {} bytes", data_xml.len());
    println!("     - Layout XML: {} bytes", layout_xml.len());
    println!("     - Colors XML: {} bytes", colors_xml.len());
    println!("     - Style XML: {} bytes", style_xml.len());
    println!("     - Total nodes: {}", enterprise_org.node_count());

    let all_text = enterprise_org.text();
    println!("     - Contains {} characters of text", all_text.len());
    println!();

    Ok(())
}
