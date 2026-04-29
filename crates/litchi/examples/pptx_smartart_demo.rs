//! SmartArt Demo Example
//!
//! This example creates a PPTX file with various SmartArt diagram types embedded in slides.
//! The SmartArt diagrams are fully functional and can be opened in Microsoft PowerPoint.
//!
//! Run with: cargo run --example pptx_smartart_demo --features ooxml

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::smartart::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== PPTX SmartArt Demo ===\n");
    println!("Creating presentation with embedded SmartArt diagrams...\n");

    let mut pkg = Package::new()?;

    {
        let pres = pkg.presentation_mut()?;

        // ====================================================================
        // Slide 1: Title Slide
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("SmartArt Diagram Types Demo");
            slide.add_text_box(
                "Demonstrating various SmartArt diagram types in PowerPoint\nCreated with Litchi",
                914400,
                3000000,
                7315200,
                1000000,
            );
        }

        // ====================================================================
        // Slide 2: List Diagram - Project Tasks
        // ====================================================================
        {
            let task_list = SmartArtBuilder::new(DiagramType::List)
                .layout_name("Basic Block List")
                .add_items(vec![
                    "Define project scope and objectives",
                    "Allocate resources and budget",
                    "Set milestones and deadlines",
                    "Monitor progress and risks",
                    "Deliver results and documentation",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&task_list)?;
            let slide = pres.add_slide()?;
            slide.set_title("List Diagram - Project Tasks");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added List Diagram (Project Tasks)");
        }

        // ====================================================================
        // Slide 3: Process Diagram - Software Development
        // ====================================================================
        {
            let dev_process = SmartArtBuilder::new(DiagramType::Process)
                .layout_name("Basic Process")
                .add_items(vec![
                    "Requirements",
                    "Design",
                    "Development",
                    "Testing",
                    "Deployment",
                    "Maintenance",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&dev_process)?;
            let slide = pres.add_slide()?;
            slide.set_title("Process Diagram - Development Lifecycle");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Process Diagram (Development Lifecycle)");
        }

        // ====================================================================
        // Slide 4: Cycle Diagram - PDCA Cycle
        // ====================================================================
        {
            let pdca_cycle = SmartArtBuilder::new(DiagramType::Cycle)
                .layout_name("Basic Cycle")
                .add_items(vec![
                    "Plan: Identify opportunities for improvement",
                    "Do: Implement changes on small scale",
                    "Check: Evaluate results and data",
                    "Act: Standardize or adjust approach",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&pdca_cycle)?;
            let slide = pres.add_slide()?;
            slide.set_title("Cycle Diagram - Continuous Improvement");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Cycle Diagram (PDCA)");
        }

        // ====================================================================
        // Slide 5: Hierarchy Diagram - Organization Chart
        // ====================================================================
        {
            let mut org_chart = SmartArt::new(DiagramType::Hierarchy);
            org_chart.layout_name = Some("Organization Chart".to_string());

            let mut ceo = DiagramNode::new("CEO - Jane Smith");
            ceo.depth = 0;

            let mut cfo = DiagramNode::new("CFO - John Doe");
            cfo.depth = 1;
            cfo.add_child(DiagramNode::new("Finance Manager"));
            cfo.add_child(DiagramNode::new("Accounting Manager"));

            let mut cto = DiagramNode::new("CTO - Alice Johnson");
            cto.depth = 1;
            cto.add_child(DiagramNode::new("Engineering Lead"));
            cto.add_child(DiagramNode::new("DevOps Lead"));
            cto.add_child(DiagramNode::new("QA Lead"));

            let mut coo = DiagramNode::new("COO - Bob Wilson");
            coo.depth = 1;
            coo.add_child(DiagramNode::new("Operations Manager"));
            coo.add_child(DiagramNode::new("HR Manager"));

            ceo.add_child(cfo);
            ceo.add_child(cto);
            ceo.add_child(coo);
            org_chart.add_node(ceo);

            let diagram_idx = pres.add_smartart_parts(&org_chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Hierarchy Diagram - Organization Chart");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Hierarchy Diagram (Org Chart)");
        }

        // ====================================================================
        // Slide 6: Relationship Diagram - Venn
        // ====================================================================
        {
            let venn = SmartArtBuilder::new(DiagramType::Relationship)
                .layout_name("Basic Venn")
                .add_items(vec![
                    "Skills: What you're good at",
                    "Passion: What you love doing",
                    "Market Need: What people will pay for",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&venn)?;
            let slide = pres.add_slide()?;
            slide.set_title("Relationship Diagram - Finding Your Niche");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Relationship Diagram (Venn)");
        }

        // ====================================================================
        // Slide 7: Matrix Diagram - Priority Matrix
        // ====================================================================
        {
            let priority_matrix = SmartArtBuilder::new(DiagramType::Matrix)
                .layout_name("Basic Matrix")
                .add_items(vec![
                    "Urgent & Important: Do First",
                    "Not Urgent & Important: Schedule",
                    "Urgent & Not Important: Delegate",
                    "Not Urgent & Not Important: Eliminate",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&priority_matrix)?;
            let slide = pres.add_slide()?;
            slide.set_title("Matrix Diagram - Eisenhower Priority Matrix");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Matrix Diagram (Priority Matrix)");
        }

        // ====================================================================
        // Slide 8: Pyramid Diagram - Maslow's Hierarchy
        // ====================================================================
        {
            let needs_pyramid = SmartArtBuilder::new(DiagramType::Pyramid)
                .layout_name("Basic Pyramid")
                .add_items(vec![
                    "Self-Actualization",
                    "Esteem Needs",
                    "Love & Belonging",
                    "Safety Needs",
                    "Physiological Needs",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&needs_pyramid)?;
            let slide = pres.add_slide()?;
            slide.set_title("Pyramid Diagram - Maslow's Hierarchy of Needs");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Pyramid Diagram (Maslow's Hierarchy)");
        }

        // ====================================================================
        // Slide 9: Process Diagram - Customer Journey
        // ====================================================================
        {
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

            let diagram_idx = pres.add_smartart_parts(&customer_journey)?;
            let slide = pres.add_slide()?;
            slide.set_title("Process Diagram - Customer Journey");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Process Diagram (Customer Journey)");
        }

        // ====================================================================
        // Slide 10: Cycle Diagram - Agile Sprint
        // ====================================================================
        {
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

            let diagram_idx = pres.add_smartart_parts(&agile_sprint)?;
            let slide = pres.add_slide()?;
            slide.set_title("Cycle Diagram - Agile Sprint Cycle");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Cycle Diagram (Agile Sprint)");
        }

        // ====================================================================
        // Slide 11: List Diagram - Technology Stack
        // ====================================================================
        {
            let tech_stack = SmartArtBuilder::new(DiagramType::List)
                .layout_name("Vertical Box List")
                .add_items(vec![
                    "Frontend: React, TypeScript, Tailwind CSS",
                    "Backend: Rust, Actix-web, PostgreSQL",
                    "Infrastructure: Docker, Kubernetes, AWS",
                    "DevOps: GitHub Actions, Terraform, Prometheus",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&tech_stack)?;
            let slide = pres.add_slide()?;
            slide.set_title("List Diagram - Technology Stack");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added List Diagram (Technology Stack)");
        }

        // ====================================================================
        // Slide 12: Pyramid Diagram - Strategic Planning
        // ====================================================================
        {
            let strategy_pyramid = SmartArtBuilder::new(DiagramType::Pyramid)
                .layout_name("Inverted Pyramid")
                .add_items(vec![
                    "Vision & Mission",
                    "Strategic Goals",
                    "Tactical Objectives",
                    "Operational Plans",
                    "Daily Activities",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&strategy_pyramid)?;
            let slide = pres.add_slide()?;
            slide.set_title("Pyramid Diagram - Strategic Planning");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 4000000);
            println!("  ✓ Added Pyramid Diagram (Strategic Planning)");
        }

        // ====================================================================
        // Slide 13: Summary
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("SmartArt Types Summary");
            slide.add_text_box(
                "SmartArt diagrams demonstrated in this presentation:\n\n\
                 • List - Non-sequential or grouped information\n\
                 • Process - Steps in a workflow or timeline\n\
                 • Cycle - Continuous or repeating processes\n\
                 • Hierarchy - Organizational structures or rankings\n\
                 • Relationship - Connections and overlapping concepts\n\
                 • Matrix - How parts relate to a whole (2x2 grids)\n\
                 • Pyramid - Proportional or hierarchical relationships\n\n\
                 Each diagram type helps visualize different kinds of information.",
                914400,
                1600000,
                7315200,
                4500000,
            );
        }
    }

    // Save the presentation
    let output_path = "smartart_demo.pptx";
    pkg.save(output_path)?;

    println!("\n=== SmartArt Demo Complete ===");
    println!("✓ Saved: {}", output_path);
    println!("\nOpen this file in Microsoft PowerPoint to verify the SmartArt diagrams.");
    println!("Total slides: 13");
    println!("Total SmartArt diagrams: 11");

    Ok(())
}
