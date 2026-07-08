//! Combined Charts and SmartArt Demo
//!
//! This example creates a comprehensive PPTX file demonstrating both chart and
//! SmartArt features together. This is useful for verifying that both features
//! work correctly in the same presentation.
//!
//! Run with: cargo run --example pptx_charts_smartart_combined --features ooxml

use litchi::ooxml::pptx::Package;
use litchi::ooxml::pptx::parts::chart::{ChartData, ChartSeries, ChartType};
use litchi::ooxml::pptx::smartart::{DiagramNode, DiagramType, SmartArt, SmartArtBuilder};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== PPTX Charts & SmartArt Combined Demo ===\n");
    println!("Creating presentation with charts and SmartArt diagrams...\n");

    let mut pkg = Package::new()?;

    {
        let pres = pkg.presentation_mut()?;

        // Set widescreen aspect ratio
        pres.set_widescreen_slide_size();

        // ====================================================================
        // Slide 1: Title Slide
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("Business Analytics Dashboard");
            slide.add_text_box(
                "Q4 2024 Performance Review\nCharts & SmartArt Integration Demo",
                914400,
                2500000,
                7315200,
                1200000,
            );
        }

        // ====================================================================
        // Slide 2: Executive Summary with Process Diagram
        // ====================================================================
        {
            let review_process = SmartArtBuilder::new(DiagramType::Process)
                .layout_name("Basic Process")
                .add_items(vec![
                    "Data Collection",
                    "Analysis",
                    "Insights",
                    "Recommendations",
                    "Action Plan",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&review_process)?;
            let slide = pres.add_slide()?;
            slide.set_title("Executive Summary - Review Process");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3500000);
            println!("  ✓ Slide 2: Executive Summary (Process Diagram)");
        }

        // ====================================================================
        // Slide 3: Revenue Performance - Column Chart
        // ====================================================================
        {
            let quarters = vec![
                "Q1".to_string(),
                "Q2".to_string(),
                "Q3".to_string(),
                "Q4".to_string(),
            ];

            let actual = ChartSeries::new("Actual Revenue")
                .with_categories(quarters.clone())
                .with_values(vec![2.8, 3.2, 3.5, 4.1]);

            let target = ChartSeries::new("Target Revenue")
                .with_categories(quarters)
                .with_values(vec![2.5, 3.0, 3.5, 4.0]);

            let chart = ChartData::new(ChartType::Column, 914400, 1600000, 7315200, 3800000)
                .with_title("Quarterly Revenue ($M)")
                .add_series(actual)
                .add_series(target)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Revenue Performance");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 3: Revenue Performance (Column Chart)");
        }

        // ====================================================================
        // Slide 4: Market Position - Pie Chart
        // ====================================================================
        {
            let segments: Vec<String> = [
                "Our Company",
                "Competitor A",
                "Competitor B",
                "Competitor C",
                "Others",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let market_share = ChartSeries::new("Market Share")
                .with_categories(segments)
                .with_values(vec![32.5, 24.8, 18.2, 12.5, 12.0]);

            let chart = ChartData::new(ChartType::Pie, 1500000, 1600000, 6000000, 3800000)
                .with_title("Market Share Distribution")
                .add_series(market_share)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Market Position");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 4: Market Position (Pie Chart)");
        }

        // ====================================================================
        // Slide 5: Organization Structure - Hierarchy Diagram
        // ====================================================================
        {
            let mut org_chart = SmartArt::new(DiagramType::Hierarchy);
            org_chart.layout_name = Some("Organization Chart".to_string());

            let mut ceo = DiagramNode::new("CEO");
            ceo.depth = 0;

            let mut sales = DiagramNode::new("VP Sales");
            sales.depth = 1;
            sales.add_child(DiagramNode::new("Regional Directors"));

            let mut product = DiagramNode::new("VP Product");
            product.depth = 1;
            product.add_child(DiagramNode::new("Product Managers"));

            let mut engineering = DiagramNode::new("VP Engineering");
            engineering.depth = 1;
            engineering.add_child(DiagramNode::new("Tech Leads"));

            ceo.add_child(sales);
            ceo.add_child(product);
            ceo.add_child(engineering);
            org_chart.add_node(ceo);

            let diagram_idx = pres.add_smartart_parts(&org_chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Leadership Team");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3800000);
            println!("  ✓ Slide 5: Leadership Team (Hierarchy Diagram)");
        }

        // ====================================================================
        // Slide 6: Growth Trends - Line Chart
        // ====================================================================
        {
            let months: Vec<String> = [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let users = ChartSeries::new("Active Users (K)")
                .with_categories(months.clone())
                .with_values(vec![
                    125.0, 132.0, 145.0, 158.0, 172.0, 189.0, 205.0, 218.0, 235.0, 252.0, 268.0,
                    285.0,
                ]);

            let revenue = ChartSeries::new("MRR ($K)")
                .with_categories(months)
                .with_values(vec![
                    85.0, 92.0, 98.0, 108.0, 118.0, 128.0, 142.0, 155.0, 168.0, 182.0, 195.0, 210.0,
                ]);

            let chart = ChartData::new(ChartType::Line, 914400, 1600000, 7315200, 3800000)
                .with_title("User Growth & Revenue Trends")
                .add_series(users)
                .add_series(revenue)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Growth Trends");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 6: Growth Trends (Line Chart)");
        }

        // ====================================================================
        // Slide 7: Strategic Priorities - Pyramid Diagram
        // ====================================================================
        {
            let priorities = SmartArtBuilder::new(DiagramType::Pyramid)
                .layout_name("Basic Pyramid")
                .add_items(vec![
                    "Market Leadership",
                    "Product Innovation",
                    "Customer Success",
                    "Operational Excellence",
                    "Team Development",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&priorities)?;
            let slide = pres.add_slide()?;
            slide.set_title("Strategic Priorities");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3800000);
            println!("  ✓ Slide 7: Strategic Priorities (Pyramid Diagram)");
        }

        // ====================================================================
        // Slide 8: Product Performance - Bar Chart
        // ====================================================================
        {
            let products: Vec<String> = [
                "Enterprise Suite",
                "Professional",
                "Team Edition",
                "Starter Pack",
                "Free Tier",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let revenue = ChartSeries::new("Revenue ($M)")
                .with_categories(products)
                .with_values(vec![4.2, 2.8, 1.9, 0.8, 0.0]);

            let chart = ChartData::new(ChartType::Bar, 914400, 1600000, 7315200, 3800000)
                .with_title("Revenue by Product Tier")
                .add_series(revenue)
                .with_legend(false);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Product Performance");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 8: Product Performance (Bar Chart)");
        }

        // ====================================================================
        // Slide 9: Development Cycle - Cycle Diagram
        // ====================================================================
        {
            let dev_cycle = SmartArtBuilder::new(DiagramType::Cycle)
                .layout_name("Basic Cycle")
                .add_items(vec![
                    "Ideation & Research",
                    "Design & Prototype",
                    "Development & Testing",
                    "Launch & Monitor",
                    "Iterate & Improve",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&dev_cycle)?;
            let slide = pres.add_slide()?;
            slide.set_title("Product Development Cycle");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3800000);
            println!("  ✓ Slide 9: Development Cycle (Cycle Diagram)");
        }

        // ====================================================================
        // Slide 10: Regional Performance - Area Chart
        // ====================================================================
        {
            let quarters: Vec<String> = ["Q1", "Q2", "Q3", "Q4"]
                .iter()
                .map(|s| s.to_string())
                .collect();

            let americas = ChartSeries::new("Americas")
                .with_categories(quarters.clone())
                .with_values(vec![1.2, 1.4, 1.5, 1.8]);

            let emea = ChartSeries::new("EMEA")
                .with_categories(quarters.clone())
                .with_values(vec![0.8, 0.9, 1.0, 1.2]);

            let apac = ChartSeries::new("APAC")
                .with_categories(quarters)
                .with_values(vec![0.5, 0.6, 0.7, 0.9]);

            let chart = ChartData::new(ChartType::Area, 914400, 1600000, 7315200, 3800000)
                .with_title("Revenue by Region ($M)")
                .add_series(americas)
                .add_series(emea)
                .add_series(apac)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Regional Performance");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 10: Regional Performance (Area Chart)");
        }

        // ====================================================================
        // Slide 11: Key Initiatives - List Diagram
        // ====================================================================
        {
            let initiatives = SmartArtBuilder::new(DiagramType::List)
                .layout_name("Basic Block List")
                .add_items(vec![
                    "Launch Enterprise v3.0 with AI features",
                    "Expand into 5 new markets in APAC",
                    "Achieve SOC 2 Type II certification",
                    "Reduce customer churn by 15%",
                    "Hire 50 new engineers globally",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&initiatives)?;
            let slide = pres.add_slide()?;
            slide.set_title("Q1 2025 Key Initiatives");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3800000);
            println!("  ✓ Slide 11: Key Initiatives (List Diagram)");
        }

        // ====================================================================
        // Slide 12: Customer Satisfaction - Doughnut Chart
        // ====================================================================
        {
            let ratings: Vec<String> = [
                "Very Satisfied",
                "Satisfied",
                "Neutral",
                "Dissatisfied",
                "Very Dissatisfied",
            ]
            .iter()
            .map(|s| s.to_string())
            .collect();

            let satisfaction = ChartSeries::new("Satisfaction %")
                .with_categories(ratings)
                .with_values(vec![45.0, 32.0, 15.0, 5.0, 3.0]);

            let chart = ChartData::new(ChartType::Doughnut, 1500000, 1600000, 6000000, 3800000)
                .with_title("Customer Satisfaction Survey Results")
                .add_series(satisfaction)
                .with_legend(true);

            let chart_idx = pres.add_chart_parts(&chart)?;
            let slide = pres.add_slide()?;
            slide.set_title("Customer Satisfaction");
            slide.add_chart_shape(chart_idx, chart.x, chart.y, chart.width, chart.height);
            println!("  ✓ Slide 12: Customer Satisfaction (Doughnut Chart)");
        }

        // ====================================================================
        // Slide 13: Decision Framework - Matrix Diagram
        // ====================================================================
        {
            let matrix = SmartArtBuilder::new(DiagramType::Matrix)
                .layout_name("Basic Matrix")
                .add_items(vec![
                    "High Impact, Low Effort: Quick Wins",
                    "High Impact, High Effort: Major Projects",
                    "Low Impact, Low Effort: Fill-ins",
                    "Low Impact, High Effort: Avoid",
                ])
                .build();

            let diagram_idx = pres.add_smartart_parts(&matrix)?;
            let slide = pres.add_slide()?;
            slide.set_title("Investment Decision Framework");
            slide.add_smartart_shape(diagram_idx, 914400, 1600000, 7315200, 3800000);
            println!("  ✓ Slide 13: Decision Framework (Matrix Diagram)");
        }

        // ====================================================================
        // Slide 14: Thank You / Q&A
        // ====================================================================
        {
            let slide = pres.add_slide()?;
            slide.set_title("Thank You");
            slide.add_text_box(
                "Questions & Discussion\n\n\
                 This presentation demonstrates:\n\
                 • 6 Chart types (Column, Pie, Line, Bar, Area, Doughnut)\n\
                 • 6 SmartArt types (Process, Hierarchy, Pyramid, Cycle, List, Matrix)\n\n\
                 Created with Litchi - High-performance Rust Office library",
                914400,
                2000000,
                7315200,
                3500000,
            );
        }
    }

    // Save the presentation
    let output_path = "charts_smartart_combined.pptx";
    pkg.save(output_path)?;

    println!("\n=== Combined Demo Complete ===");
    println!("✓ Saved: {}", output_path);
    println!("\nOpen this file in Microsoft PowerPoint to verify:");
    println!("  - Total slides: 14");
    println!("  - Charts: 6 (Column, Pie, Line, Bar, Area, Doughnut)");
    println!("  - SmartArt: 6 (Process, Hierarchy, Pyramid, Cycle, List, Matrix)");

    Ok(())
}
