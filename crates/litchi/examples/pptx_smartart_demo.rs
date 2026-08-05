//! Typed DrawingML diagram and PresentationML package demonstration.
//!
//! SmartArt is modeled independently from the PPTX package graph. This
//! example keeps the original eleven diagram cases, validates their typed
//! models and canonical diagram-part XML, then writes and reads a PPTX summary
//! deck through the standalone `litchi-pptx` facade.
//!
//! Run with: `cargo run --example pptx_smartart_demo --features ooxml`

use litchi_pptx::Package;
use litchi_pptx::shape::diagram::{
    Builder, Graphic, Kind, Node, colors_xml, data_xml, drawing_xml, layout_xml, quickstyle_xml,
};

const X: i64 = 914_400;
const Y: i64 = 1_600_000;
const WIDTH: i64 = 7_315_200;
const HEIGHT: i64 = 4_000_000;

fn diagram(kind: Kind, layout: &str, items: &[&str]) -> Graphic {
    Builder::new(kind)
        .layout_name(layout)
        .add_items(items.iter().copied())
        .build()
}

fn hierarchy() -> Graphic {
    let mut ceo = Node::new("CEO - Jane Smith");

    let mut cfo = Node::new("CFO - John Doe");
    cfo.depth = 1;
    cfo.add_child(Node::new("Finance Manager"));
    cfo.add_child(Node::new("Accounting Manager"));

    let mut cto = Node::new("CTO - Alice Johnson");
    cto.depth = 1;
    cto.add_child(Node::new("Engineering Lead"));
    cto.add_child(Node::new("DevOps Lead"));
    cto.add_child(Node::new("QA Lead"));

    let mut coo = Node::new("COO - Bob Wilson");
    coo.depth = 1;
    coo.add_child(Node::new("Operations Manager"));
    coo.add_child(Node::new("HR Manager"));

    ceo.add_child(cfo);
    ceo.add_child(cto);
    ceo.add_child(coo);

    let mut value = Graphic::new(Kind::Hierarchy);
    value.layout_name = Some("Organization Chart".to_owned());
    value.add_node(ceo);
    value
}

fn part_bytes(value: &Graphic) -> usize {
    data_xml(value).len()
        + layout_xml(value).len()
        + drawing_xml(value, X, Y, WIDTH, HEIGHT).len()
        + colors_xml().len()
        + quickstyle_xml().len()
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let diagrams = vec![
        (
            "List Diagram - Project Tasks",
            diagram(
                Kind::List,
                "Basic Block List",
                &[
                    "Define project scope and objectives",
                    "Allocate resources and budget",
                    "Set milestones and deadlines",
                    "Monitor progress and risks",
                    "Deliver results and documentation",
                ],
            ),
        ),
        (
            "Process Diagram - Development Lifecycle",
            diagram(
                Kind::Process,
                "Basic Process",
                &[
                    "Requirements",
                    "Design",
                    "Development",
                    "Testing",
                    "Deployment",
                    "Maintenance",
                ],
            ),
        ),
        (
            "Cycle Diagram - Continuous Improvement",
            diagram(
                Kind::Cycle,
                "Basic Cycle",
                &[
                    "Plan: Identify opportunities for improvement",
                    "Do: Implement changes on small scale",
                    "Check: Evaluate results and data",
                    "Act: Standardize or adjust approach",
                ],
            ),
        ),
        ("Hierarchy Diagram - Organization Chart", hierarchy()),
        (
            "Relationship Diagram - Finding Your Niche",
            diagram(
                Kind::Relationship,
                "Basic Venn",
                &[
                    "Skills: What you're good at",
                    "Passion: What you love doing",
                    "Market Need: What people will pay for",
                ],
            ),
        ),
        (
            "Matrix Diagram - Eisenhower Priority Matrix",
            diagram(
                Kind::Matrix,
                "Basic Matrix",
                &[
                    "Urgent & Important: Do First",
                    "Not Urgent & Important: Schedule",
                    "Urgent & Not Important: Delegate",
                    "Not Urgent & Not Important: Eliminate",
                ],
            ),
        ),
        (
            "Pyramid Diagram - Maslow's Hierarchy of Needs",
            diagram(
                Kind::Pyramid,
                "Basic Pyramid",
                &[
                    "Self-Actualization",
                    "Esteem Needs",
                    "Love & Belonging",
                    "Safety Needs",
                    "Physiological Needs",
                ],
            ),
        ),
        (
            "Process Diagram - Customer Journey",
            diagram(
                Kind::Process,
                "Continuous Block Process",
                &[
                    "Awareness",
                    "Consideration",
                    "Decision",
                    "Purchase",
                    "Retention",
                    "Advocacy",
                ],
            ),
        ),
        (
            "Cycle Diagram - Agile Sprint Cycle",
            diagram(
                Kind::Cycle,
                "Circular Cycle",
                &[
                    "Sprint Planning",
                    "Daily Standups",
                    "Development",
                    "Testing & Review",
                    "Sprint Retrospective",
                ],
            ),
        ),
        (
            "List Diagram - Technology Stack",
            diagram(
                Kind::List,
                "Vertical Box List",
                &[
                    "Frontend: React, TypeScript, Tailwind CSS",
                    "Backend: Rust, Actix-web, PostgreSQL",
                    "Infrastructure: Docker, Kubernetes, AWS",
                    "DevOps: GitHub Actions, Terraform, Prometheus",
                ],
            ),
        ),
        (
            "Pyramid Diagram - Strategic Planning",
            diagram(
                Kind::Pyramid,
                "Inverted Pyramid",
                &[
                    "Vision & Mission",
                    "Strategic Goals",
                    "Tactical Objectives",
                    "Operational Plans",
                    "Daily Activities",
                ],
            ),
        ),
    ];

    println!("=== Typed PPTX SmartArt Demo ===\n");
    for (index, (title, value)) in diagrams.iter().enumerate() {
        println!(
            "  diagram {}: {} ({:?}, {} nodes, {} canonical XML bytes)",
            index + 1,
            title,
            value.diagram_type,
            value.node_count(),
            part_bytes(value),
        );
    }

    let mut package = Package::new()?;
    {
        let presentation = package.presentation_mut()?;
        presentation.set_widescreen_slide_size();

        let title = presentation.add_slide()?;
        title.set_title("SmartArt Diagram Types Demo");
        title.add_text_box(
            "Typed DrawingML diagrams\nStandalone PresentationML package read/write demonstration",
            X,
            2_500_000,
            WIDTH,
            1_200_000,
        );

        for (index, (title, value)) in diagrams.iter().enumerate() {
            let slide = presentation.add_slide()?;
            slide.set_title(title);
            slide.add_text_box(
                &format!(
                    "Typed diagram model: {:?}\nLayout: {}\nNodes: {}\nCanonical diagram-part XML: {} bytes\n\nThe model and codecs are shared DrawingML semantics; this facade currently demonstrates package authoring with an inspectable summary slide.",
                    value.diagram_type,
                    value.layout_name.as_deref().unwrap_or("default"),
                    value.node_count(),
                    part_bytes(value),
                ),
                X,
                Y,
                WIDTH,
                HEIGHT,
            );
            println!("  ✓ summary slide {}: {}", index + 2, title);
        }

        let summary = presentation.add_slide()?;
        summary.set_title("SmartArt Types Summary");
        summary.add_text_box(
            "List · Process · Cycle · Hierarchy · Relationship · Matrix · Pyramid\n\nEach case above is represented by the typed diagram model and its canonical DrawingML data, layout, drawing, color, and quick-style codecs.",
            X,
            2_000_000,
            WIDTH,
            2_000_000,
        );
    }

    let bytes = package.to_bytes()?;
    let output_path = "smartart_demo.pptx";
    std::fs::write(output_path, &bytes)?;

    let reopened = Package::from_bytes(&bytes)?;
    let presentation = reopened.presentation()?;
    println!("\nRead-back slide count: {}", presentation.slide_count()?);
    println!("Read-back text bytes: {}", presentation.text()?.len());
    println!("✓ Saved: {output_path}");
    Ok(())
}
