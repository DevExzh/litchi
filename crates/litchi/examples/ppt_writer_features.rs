//! Example demonstrating all new PPT writer features
//!
//! This example showcases:
//! - Pictures/Images
//! - Text formatting (bold, italic, font size, colors)
//! - Shape styling (fill colors, line styles)
//! - More shape types (lines, ellipses, arrows)
//! - Hyperlinks
//! - Notes slides
//!
//! Run with: cargo run --example ppt_writer_features

use litchi::ole::ppt::writer::{
    FillStyle, FontEntity, Hyperlink, LineStyleConfig, NotesPage, Paragraph, PptWriter,
    ShadowStyle, ShapeColor, ShapeStyle, ShapeType, TextRun,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("=== PPT Writer Feature Demo ===\n");

    // Create output directory
    std::fs::create_dir_all("output")?;

    // Generate individual examples
    create_shapes_demo()?;
    create_text_formatting_demo()?;
    create_styled_shapes_demo()?;
    create_lines_arrows_demo()?;
    create_pictures_demo()?;
    create_hyperlinks_demo()?;
    create_notes_demo()?;
    create_comprehensive_demo()?;

    println!("\n✅ All demos created successfully!");
    println!("   Open the PPT files in output/ folder to verify.");

    Ok(())
}

/// Demo 1: Basic shape types
fn create_shapes_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating shapes demo...");
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;

    // Title
    writer.add_textbox(slide, 50, 30, 600, 40, "Shape Types Demo")?;

    // Rectangle
    writer.add_rectangle(slide, 50, 100, 150, 100)?;
    writer.add_textbox(slide, 50, 210, 150, 30, "Rectangle")?;

    // Ellipse
    writer.add_ellipse(slide, 250, 100, 150, 100)?;
    writer.add_textbox(slide, 250, 210, 150, 30, "Ellipse")?;

    // Another rectangle
    writer.add_rectangle(slide, 450, 100, 150, 100)?;
    writer.add_textbox(slide, 450, 210, 150, 30, "Rectangle 2")?;

    writer.save("output/01_shapes.ppt")?;
    println!("  ✓ output/01_shapes.ppt");
    Ok(())
}

/// Demo 2: Text formatting with rich text
fn create_text_formatting_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating text formatting demo...");
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;

    // Title
    writer.add_textbox(slide, 50, 30, 600, 40, "Text Formatting Demo")?;

    // Bold text
    let bold_para = Paragraph::with_runs(vec![
        TextRun::new("This text is ").size(18),
        TextRun::new("BOLD").bold().size(18),
        TextRun::new(" and this is normal.").size(18),
    ]);
    writer.add_rich_textbox(slide, 50, 100, 600, 40, vec![bold_para])?;

    // Italic text
    let italic_para = Paragraph::with_runs(vec![
        TextRun::new("This text is ").size(18),
        TextRun::new("italic").italic().size(18),
        TextRun::new(" style.").size(18),
    ]);
    writer.add_rich_textbox(slide, 50, 150, 600, 40, vec![italic_para])?;

    // Bold italic
    let bold_italic_para = Paragraph::with_runs(vec![
        TextRun::new("Combined: ").size(18),
        TextRun::new("bold and italic").bold().italic().size(18),
        TextRun::new(" together!").size(18),
    ]);
    writer.add_rich_textbox(slide, 50, 200, 600, 40, vec![bold_italic_para])?;

    // Different font sizes
    let sizes_para = Paragraph::with_runs(vec![
        TextRun::new("Small ").size(12),
        TextRun::new("Medium ").size(18),
        TextRun::new("Large ").size(24),
        TextRun::new("HUGE").size(36),
    ]);
    writer.add_rich_textbox(slide, 50, 260, 600, 60, vec![sizes_para])?;

    // Colored text
    let colors_para = Paragraph::with_runs(vec![
        TextRun::new("Red ").size(18).color_rgb(255, 0, 0),
        TextRun::new("Green ").size(18).color_rgb(0, 255, 0),
        TextRun::new("Blue ").size(18).color_rgb(0, 0, 255),
        TextRun::new("Orange").size(18).color_hex(0xFF8800),
    ]);
    writer.add_rich_textbox(slide, 50, 340, 600, 40, vec![colors_para])?;

    // Underlined text
    let underline_para = Paragraph::with_runs(vec![
        TextRun::new("This text has ").size(18),
        TextRun::new("underline").underline().size(18),
        TextRun::new(" formatting.").size(18),
    ]);
    writer.add_rich_textbox(slide, 50, 400, 600, 40, vec![underline_para])?;

    writer.save("output/02_text_formatting.ppt")?;
    println!("  ✓ output/02_text_formatting.ppt");
    Ok(())
}

/// Demo 3: Shape styling (fills, lines, shadows)
fn create_styled_shapes_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating styled shapes demo...");
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;

    // Title
    writer.add_textbox(slide, 50, 20, 600, 40, "Shape Styling Demo")?;

    // Solid red fill
    let red_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(255, 0, 0))
        .with_line(LineStyleConfig::none());
    writer.add_styled_shape(slide, ShapeType::Rectangle, 50, 80, 120, 80, red_style)?;
    writer.add_textbox(slide, 50, 170, 120, 25, "Red Fill")?;

    // Blue fill with black border
    let blue_bordered_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(0, 100, 200))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            2.0,
        ));
    writer.add_styled_shape(
        slide,
        ShapeType::Rectangle,
        200,
        80,
        120,
        80,
        blue_bordered_style,
    )?;
    writer.add_textbox(slide, 200, 170, 120, 25, "Blue + Border")?;

    // Green ellipse
    let green_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(0, 180, 0))
        .with_line(LineStyleConfig::none());
    writer.add_styled_shape(slide, ShapeType::Ellipse, 350, 80, 120, 80, green_style)?;
    writer.add_textbox(slide, 350, 170, 120, 25, "Green Ellipse")?;

    // Yellow with thick dashed border
    let mut dashed_line =
        LineStyleConfig::with_color_and_width(ShapeColor::rgb(100, 100, 100), 3.0);
    dashed_line.dash = litchi::ole::ppt::writer::shape_style::LineDashStyle::Dash;
    let yellow_dashed_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(255, 255, 0))
        .with_line(dashed_line);
    writer.add_styled_shape(
        slide,
        ShapeType::Rectangle,
        500,
        80,
        120,
        80,
        yellow_dashed_style,
    )?;
    writer.add_textbox(slide, 500, 170, 120, 25, "Dashed Border")?;

    // No fill, only border
    let no_fill_style = ShapeStyle::no_fill();
    writer.add_styled_shape(slide, ShapeType::Rectangle, 50, 220, 120, 80, no_fill_style)?;
    writer.add_textbox(slide, 50, 310, 120, 25, "No Fill")?;

    // With shadow
    let shadow_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(200, 200, 255))
        .with_line(LineStyleConfig::default_line())
        .with_shadow(ShadowStyle::drop_shadow());
    writer.add_styled_shape(slide, ShapeType::Rectangle, 200, 220, 120, 80, shadow_style)?;
    writer.add_textbox(slide, 200, 310, 120, 25, "With Shadow")?;

    // Gradient fill
    let gradient_style = ShapeStyle::new()
        .with_fill(FillStyle::gradient(
            ShapeColor::rgb(0, 100, 200),
            ShapeColor::rgb(200, 200, 255),
            45,
        ))
        .with_line(LineStyleConfig::none());
    writer.add_styled_shape(
        slide,
        ShapeType::Rectangle,
        350,
        220,
        120,
        80,
        gradient_style,
    )?;
    writer.add_textbox(slide, 350, 310, 120, 25, "Gradient")?;

    // Semi-transparent
    let transparent_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(255, 0, 255).with_opacity(50))
        .with_line(LineStyleConfig::default_line());
    writer.add_styled_shape(
        slide,
        ShapeType::Ellipse,
        500,
        220,
        120,
        80,
        transparent_style,
    )?;
    writer.add_textbox(slide, 500, 310, 120, 25, "50% Opacity")?;

    writer.save("output/03_styled_shapes.ppt")?;
    println!("  ✓ output/03_styled_shapes.ppt");
    Ok(())
}

/// Demo 4: Lines and arrows
fn create_lines_arrows_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating lines and arrows demo...");
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;

    // Title
    writer.add_textbox(slide, 50, 20, 600, 40, "Lines and Arrows Demo")?;

    // Simple horizontal line
    writer.add_line(slide, 50, 100, 250, 100)?;
    writer.add_textbox(slide, 50, 110, 200, 25, "Horizontal Line")?;

    // Simple vertical line
    writer.add_line(slide, 300, 80, 300, 180)?;
    writer.add_textbox(slide, 310, 120, 150, 25, "Vertical Line")?;

    // Diagonal line
    writer.add_line(slide, 450, 80, 600, 180)?;
    writer.add_textbox(slide, 500, 190, 100, 25, "Diagonal")?;

    // Arrow pointing right
    writer.add_arrow_line(slide, 50, 250, 250, 250)?;
    writer.add_textbox(slide, 50, 260, 200, 25, "Arrow Right")?;

    // Arrow pointing left
    writer.add_arrow_line(slide, 450, 250, 300, 250)?;
    writer.add_textbox(slide, 350, 260, 150, 25, "Arrow Left")?;

    // Arrow pointing down
    writer.add_arrow_line(slide, 550, 220, 550, 350)?;
    writer.add_textbox(slide, 560, 280, 100, 25, "Down")?;

    // Arrow pointing up
    writer.add_arrow_line(slide, 600, 350, 600, 220)?;
    writer.add_textbox(slide, 610, 280, 100, 25, "Up")?;

    // Diagonal arrow
    writer.add_arrow_line(slide, 50, 350, 200, 450)?;
    writer.add_textbox(slide, 100, 460, 150, 25, "Diagonal Arrow")?;

    // Create a flowchart-like diagram with boxes and arrows
    writer.add_rectangle(slide, 300, 350, 80, 50)?;
    writer.add_textbox(slide, 305, 365, 70, 20, "Start")?;

    writer.add_arrow_line(slide, 380, 375, 430, 375)?;

    writer.add_rectangle(slide, 430, 350, 80, 50)?;
    writer.add_textbox(slide, 435, 365, 70, 20, "Process")?;

    writer.add_arrow_line(slide, 510, 375, 560, 375)?;

    writer.add_rectangle(slide, 560, 350, 80, 50)?;
    writer.add_textbox(slide, 565, 365, 70, 20, "End")?;

    writer.save("output/04_lines_arrows.ppt")?;
    println!("  ✓ output/04_lines_arrows.ppt");
    Ok(())
}

/// Demo 5: Pictures/Images
fn create_pictures_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating pictures demo...");
    let mut writer = PptWriter::new();

    let slide = writer.add_slide()?;
    writer.add_textbox(slide, 50, 20, 600, 40, "Pictures Demo")?;

    // Load the litchi logo
    let logo_path = std::path::Path::new("media/litchi_logo.png");
    if logo_path.exists() {
        let image_data = std::fs::read(logo_path)?;

        // Add picture at different sizes
        writer.add_picture(slide, 50, 80, 150, 150, image_data.clone())?;
        writer.add_textbox(slide, 50, 240, 150, 25, "Original Size")?;

        writer.add_picture(slide, 250, 80, 100, 100, image_data.clone())?;
        writer.add_textbox(slide, 250, 190, 100, 25, "Small")?;

        writer.add_picture(slide, 400, 80, 200, 200, image_data)?;
        writer.add_textbox(slide, 400, 290, 200, 25, "Large")?;
    } else {
        // Create placeholder rectangles if image not found
        writer.add_textbox(
            slide,
            50,
            100,
            600,
            100,
            "(Image file not found at media/litchi_logo.png)\n\n\
             To test pictures, place a PNG file there.",
        )?;
    }

    // Add second slide with different image layout
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 50, 20, 600, 40, "Pictures with Shapes")?;

    // Load gopher image if available
    let gopher_path = std::path::Path::new("soapberry-zip/assets/gophercolor16x16.png");
    if gopher_path.exists() {
        let gopher_data = std::fs::read(gopher_path)?;

        // Grid of small images
        for row in 0..3 {
            for col in 0..5 {
                let x = 50 + col * 70;
                let y = 80 + row * 70;
                writer.add_picture(slide2, x, y, 60, 60, gopher_data.clone())?;
            }
        }
        writer.add_textbox(
            slide2,
            50,
            300,
            400,
            30,
            "Grid of 16x16 gopher images scaled to 60x60",
        )?;
    } else {
        writer.add_textbox(slide2, 50, 100, 600, 50, "(Gopher image not found)")?;
    }

    println!("  Picture count: {}", writer.picture_count());
    writer.save("output/05_pictures.ppt")?;
    println!("  ✓ output/05_pictures.ppt");
    Ok(())
}

/// Demo 6: Hyperlinks
fn create_hyperlinks_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating hyperlinks demo...");
    let mut writer = PptWriter::new();

    // Slide 1: External links
    let slide1 = writer.add_slide()?;
    writer.add_textbox(slide1, 50, 30, 600, 40, "Hyperlinks Demo - External Links")?;

    // Add URL hyperlink
    let url_link_id =
        writer.add_hyperlink(Hyperlink::url("https://github.com").with_display_text("GitHub"));
    writer.add_rectangle(slide1, 50, 100, 200, 50)?;
    writer.add_textbox(slide1, 55, 115, 190, 30, "Click: GitHub")?;
    writer.set_last_shape_hyperlink(slide1, url_link_id)?;

    // Add another URL
    let rust_link_id = writer.add_hyperlink(
        Hyperlink::url("https://www.rust-lang.org").with_display_text("Rust Language"),
    );
    writer.add_rectangle(slide1, 50, 180, 200, 50)?;
    writer.add_textbox(slide1, 55, 195, 190, 30, "Click: Rust Lang")?;
    writer.set_last_shape_hyperlink(slide1, rust_link_id)?;

    // File link
    let file_link_id = writer.add_hyperlink(
        Hyperlink::file("C:\\Documents\\report.pdf").with_display_text("Open Report"),
    );
    writer.add_rectangle(slide1, 50, 260, 200, 50)?;
    writer.add_textbox(slide1, 55, 275, 190, 30, "File Link")?;
    writer.set_last_shape_hyperlink(slide1, file_link_id)?;

    // Slide 2: Navigation links
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 50, 30, 600, 40, "Hyperlinks Demo - Navigation")?;

    // Next slide link
    let next_link_id = writer.add_hyperlink(Hyperlink::next_slide());
    writer.add_rectangle(slide2, 50, 100, 150, 50)?;
    writer.add_textbox(slide2, 55, 115, 140, 30, "Next Slide")?;
    writer.set_last_shape_hyperlink(slide2, next_link_id)?;

    // Previous slide link
    let prev_link_id = writer.add_hyperlink(Hyperlink::prev_slide());
    writer.add_rectangle(slide2, 220, 100, 150, 50)?;
    writer.add_textbox(slide2, 225, 115, 140, 30, "Previous Slide")?;
    writer.set_last_shape_hyperlink(slide2, prev_link_id)?;

    // Jump to specific slide
    let slide_link_id = writer.add_hyperlink(Hyperlink::slide(1));
    writer.add_rectangle(slide2, 390, 100, 150, 50)?;
    writer.add_textbox(slide2, 395, 115, 140, 30, "Go to Slide 1")?;
    writer.set_last_shape_hyperlink(slide2, slide_link_id)?;

    // Slide 3: Target slide
    let slide3 = writer.add_slide()?;
    writer.add_textbox(
        slide3,
        50,
        200,
        600,
        60,
        "This is Slide 3 - Navigation Target",
    )?;

    let back_link_id = writer.add_hyperlink(Hyperlink::slide(2));
    writer.add_rectangle(slide3, 250, 300, 200, 50)?;
    writer.add_textbox(slide3, 270, 315, 160, 30, "Back to Slide 2")?;
    writer.set_last_shape_hyperlink(slide3, back_link_id)?;

    writer.save("output/06_hyperlinks.ppt")?;
    println!("  ✓ output/06_hyperlinks.ppt");
    Ok(())
}

/// Demo 6: Notes slides
fn create_notes_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating notes demo...");
    let mut writer = PptWriter::new();

    // Slide 1 with simple notes
    let slide1 = writer.add_slide()?;
    writer.add_textbox(slide1, 50, 50, 600, 100, "Introduction")?;
    writer.add_textbox(
        slide1,
        50,
        180,
        600,
        200,
        "This slide demonstrates speaker notes.\n\nLook at Notes view to see them!",
    )?;
    writer.set_slide_notes(
        slide1,
        "SPEAKER NOTES for Introduction:\n\n\
         - Welcome the audience\n\
         - Introduce yourself\n\
         - Give overview of presentation\n\
         - Mention that there will be Q&A at the end",
    )?;

    // Slide 2 with rich notes
    let slide2 = writer.add_slide()?;
    writer.add_textbox(slide2, 50, 50, 600, 100, "Main Content")?;
    writer.add_textbox(slide2, 50, 180, 600, 200, "Key points go here...")?;

    // Use rich notes page
    let notes_page = NotesPage::new(2) // slide_id_ref will be set during save
        .with_text(
            "Detailed speaker notes for the main content slide:\n\n\
             1. Explain the first key point in detail\n\
             2. Use examples from real-world scenarios\n\
             3. Pause for questions\n\
             4. Transition to the next topic smoothly\n\n\
             TIMING: This slide should take about 5 minutes.",
        )
        .with_footer("Confidential - Internal Use Only");
    writer.set_notes_page(slide2, notes_page)?;

    // Slide 3 with no notes
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 50, 50, 600, 100, "Conclusion")?;
    writer.add_textbox(slide3, 50, 180, 600, 200, "Thank you for your attention!")?;
    // No notes on this slide - intentional

    // Slide 4 with extensive notes
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 50, 50, 600, 100, "Q&A Session")?;
    writer.set_slide_notes(
        slide4,
        "Q&A PREPARATION:\n\n\
         Anticipated questions:\n\
         Q: What is the timeline?\n\
         A: We expect completion by Q4 2024.\n\n\
         Q: What are the costs?\n\
         A: See appendix for detailed breakdown.\n\n\
         Q: Who is responsible?\n\
         A: The core team with support from IT.\n\n\
         FALLBACK: If no questions, summarize key takeaways.",
    )?;

    writer.save("output/07_notes.ppt")?;
    println!("  ✓ output/07_notes.ppt");
    Ok(())
}

/// Demo 7: Comprehensive demo with all features
fn create_comprehensive_demo() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating comprehensive demo...");
    let mut writer = PptWriter::new();

    // Add custom font
    let _times_font = writer.add_font(FontEntity::times_new_roman());

    // === SLIDE 1: Title Slide ===
    let slide1 = writer.add_slide()?;

    // Title with large bold text
    let title = Paragraph::with_runs(vec![
        TextRun::new("PPT Writer Features")
            .bold()
            .size(44)
            .color_rgb(0, 51, 102),
    ]);
    writer.add_rich_textbox(slide1, 50, 150, 600, 80, vec![title.center()])?;

    // Subtitle
    let subtitle = Paragraph::with_runs(vec![
        TextRun::new("Complete Feature Demonstration")
            .italic()
            .size(24)
            .color_rgb(100, 100, 100),
    ]);
    writer.add_rich_textbox(slide1, 50, 250, 600, 50, vec![subtitle.center()])?;

    // Decorative line
    writer.add_line(slide1, 150, 320, 550, 320)?;

    writer.set_slide_notes(
        slide1,
        "Title slide - introduce the comprehensive feature demonstration.",
    )?;

    // === SLIDE 2: Text Formatting Showcase ===
    let slide2 = writer.add_slide()?;

    writer.add_textbox(slide2, 50, 30, 600, 40, "Text Formatting")?;

    // Mixed formatting paragraph
    let mixed = Paragraph::with_runs(vec![
        TextRun::new("This paragraph demonstrates ").size(16),
        TextRun::new("bold").bold().size(16),
        TextRun::new(", ").size(16),
        TextRun::new("italic").italic().size(16),
        TextRun::new(", ").size(16),
        TextRun::new("underline").underline().size(16),
        TextRun::new(", and ").size(16),
        TextRun::new("colored").size(16).color_rgb(255, 0, 100),
        TextRun::new(" text all in one!").size(16),
    ]);
    writer.add_rich_textbox(slide2, 50, 90, 600, 50, vec![mixed])?;

    // Size progression
    let sizes = Paragraph::with_runs(vec![
        TextRun::new("10pt ").size(10),
        TextRun::new("14pt ").size(14),
        TextRun::new("18pt ").size(18),
        TextRun::new("24pt ").size(24),
        TextRun::new("32pt").size(32),
    ]);
    writer.add_rich_textbox(slide2, 50, 160, 600, 60, vec![sizes])?;

    // Rainbow text
    let rainbow = Paragraph::with_runs(vec![
        TextRun::new("R").size(28).bold().color_rgb(255, 0, 0),
        TextRun::new("A").size(28).bold().color_rgb(255, 127, 0),
        TextRun::new("I").size(28).bold().color_rgb(255, 255, 0),
        TextRun::new("N").size(28).bold().color_rgb(0, 255, 0),
        TextRun::new("B").size(28).bold().color_rgb(0, 0, 255),
        TextRun::new("O").size(28).bold().color_rgb(75, 0, 130),
        TextRun::new("W").size(28).bold().color_rgb(148, 0, 211),
    ]);
    writer.add_rich_textbox(slide2, 50, 250, 600, 50, vec![rainbow.center()])?;

    // === SLIDE 3: Shapes Gallery ===
    let slide3 = writer.add_slide()?;
    writer.add_textbox(slide3, 50, 20, 600, 35, "Shapes Gallery")?;

    // Row 1: Basic shapes
    let styles = [
        FillStyle::solid_rgb(255, 100, 100),
        FillStyle::solid_rgb(100, 255, 100),
        FillStyle::solid_rgb(100, 100, 255),
        FillStyle::solid_rgb(255, 255, 100),
    ];

    for (i, fill) in styles.iter().enumerate() {
        let x = 50 + (i as i32) * 160;
        let style = ShapeStyle::new().with_fill(fill.clone()).with_line(
            LineStyleConfig::with_color_and_width(ShapeColor::BLACK, 1.5),
        );
        writer.add_styled_shape(slide3, ShapeType::Rectangle, x, 70, 130, 80, style)?;
    }

    // Row 2: Ellipses
    for (i, fill) in styles.iter().enumerate() {
        let x = 50 + (i as i32) * 160;
        let style = ShapeStyle::new()
            .with_fill(fill.clone())
            .with_line(LineStyleConfig::none());
        writer.add_styled_shape(slide3, ShapeType::Ellipse, x, 180, 130, 80, style)?;
    }

    // Row 3: Mixed with effects
    let shadow_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(200, 150, 255))
        .with_shadow(ShadowStyle::drop_shadow());
    writer.add_styled_shape(slide3, ShapeType::Rectangle, 50, 290, 130, 80, shadow_style)?;

    let gradient_style = ShapeStyle::new().with_fill(FillStyle::gradient(
        ShapeColor::rgb(255, 200, 200),
        ShapeColor::rgb(200, 200, 255),
        90,
    ));
    writer.add_styled_shape(
        slide3,
        ShapeType::Rectangle,
        210,
        290,
        130,
        80,
        gradient_style,
    )?;

    let transparent_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(0, 200, 200).with_opacity(60))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            2.0,
        ));
    writer.add_styled_shape(
        slide3,
        ShapeType::Ellipse,
        370,
        290,
        130,
        80,
        transparent_style,
    )?;

    let no_fill_style = ShapeStyle::no_fill();
    writer.add_styled_shape(
        slide3,
        ShapeType::Rectangle,
        530,
        290,
        130,
        80,
        no_fill_style,
    )?;

    writer.set_slide_notes(
        slide3,
        "Shapes Gallery:\n\
         - Top row: Colored rectangles with borders\n\
         - Middle row: Colored ellipses without borders\n\
         - Bottom row: Effects (shadow, gradient, transparency, outline only)",
    )?;

    // === SLIDE 4: Lines and Connectors ===
    let slide4 = writer.add_slide()?;
    writer.add_textbox(slide4, 50, 20, 600, 35, "Lines, Arrows & Connectors")?;

    // Process flow diagram
    let box_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(220, 235, 250))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 100, 180),
            1.5,
        ));

    // Box 1
    writer.add_styled_shape(
        slide4,
        ShapeType::Rectangle,
        50,
        100,
        100,
        60,
        box_style.clone(),
    )?;
    writer.add_textbox(slide4, 55, 120, 90, 30, "Input")?;

    // Arrow 1->2
    writer.add_arrow_line(slide4, 150, 130, 200, 130)?;

    // Box 2
    writer.add_styled_shape(
        slide4,
        ShapeType::Rectangle,
        200,
        100,
        100,
        60,
        box_style.clone(),
    )?;
    writer.add_textbox(slide4, 205, 120, 90, 30, "Process")?;

    // Arrow 2->3
    writer.add_arrow_line(slide4, 300, 130, 350, 130)?;

    // Box 3
    writer.add_styled_shape(
        slide4,
        ShapeType::Rectangle,
        350,
        100,
        100,
        60,
        box_style.clone(),
    )?;
    writer.add_textbox(slide4, 355, 120, 90, 30, "Validate")?;

    // Arrow 3->4
    writer.add_arrow_line(slide4, 450, 130, 500, 130)?;

    // Box 4
    writer.add_styled_shape(slide4, ShapeType::Rectangle, 500, 100, 100, 60, box_style)?;
    writer.add_textbox(slide4, 505, 120, 90, 30, "Output")?;

    // Various line styles
    writer.add_textbox(slide4, 50, 200, 200, 25, "Line Styles:")?;
    writer.add_line(slide4, 50, 240, 200, 240)?;
    writer.add_textbox(slide4, 220, 230, 100, 25, "Solid")?;

    writer.add_arrow_line(slide4, 50, 280, 200, 280)?;
    writer.add_textbox(slide4, 220, 270, 100, 25, "Arrow")?;

    // Diagonal arrows
    writer.add_textbox(slide4, 350, 200, 200, 25, "Diagonal Arrows:")?;
    writer.add_arrow_line(slide4, 350, 250, 450, 330)?;
    writer.add_arrow_line(slide4, 450, 250, 350, 330)?;
    writer.add_arrow_line(slide4, 500, 250, 600, 250)?;
    writer.add_arrow_line(slide4, 550, 280, 550, 350)?;

    // === SLIDE 5: Interactive Elements ===
    let slide5 = writer.add_slide()?;
    writer.add_textbox(slide5, 50, 30, 600, 40, "Interactive Elements (Hyperlinks)")?;

    // Web links
    let github_id = writer.add_hyperlink(Hyperlink::url("https://github.com"));
    let button_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(0, 120, 215))
        .with_line(LineStyleConfig::none());
    writer.add_styled_shape(
        slide5,
        ShapeType::Rectangle,
        50,
        100,
        180,
        45,
        button_style.clone(),
    )?;
    let github_text = Paragraph::with_runs(vec![
        TextRun::new("Visit GitHub")
            .bold()
            .size(16)
            .color_rgb(255, 255, 255),
    ]);
    writer.add_rich_textbox(slide5, 55, 110, 170, 30, vec![github_text.center()])?;
    writer.set_last_shape_hyperlink(slide5, github_id)?;

    let rust_id = writer.add_hyperlink(Hyperlink::url("https://www.rust-lang.org"));
    writer.add_styled_shape(
        slide5,
        ShapeType::Rectangle,
        250,
        100,
        180,
        45,
        button_style,
    )?;
    let rust_text = Paragraph::with_runs(vec![
        TextRun::new("Rust Language")
            .bold()
            .size(16)
            .color_rgb(255, 255, 255),
    ]);
    writer.add_rich_textbox(slide5, 255, 110, 170, 30, vec![rust_text.center()])?;
    writer.set_last_shape_hyperlink(slide5, rust_id)?;

    // Navigation links
    writer.add_textbox(slide5, 50, 200, 300, 30, "Slide Navigation:")?;

    let nav_style = ShapeStyle::new()
        .with_fill(FillStyle::solid_rgb(100, 100, 100))
        .with_line(LineStyleConfig::none());

    let prev_id = writer.add_hyperlink(Hyperlink::prev_slide());
    writer.add_styled_shape(
        slide5,
        ShapeType::Rectangle,
        50,
        250,
        120,
        40,
        nav_style.clone(),
    )?;
    let prev_text = Paragraph::with_runs(vec![
        TextRun::new("◀ Previous").size(14).color_rgb(255, 255, 255),
    ]);
    writer.add_rich_textbox(slide5, 55, 258, 110, 25, vec![prev_text.center()])?;
    writer.set_last_shape_hyperlink(slide5, prev_id)?;

    let first_id = writer.add_hyperlink(Hyperlink::slide(1));
    writer.add_styled_shape(
        slide5,
        ShapeType::Rectangle,
        190,
        250,
        120,
        40,
        nav_style.clone(),
    )?;
    let first_text = Paragraph::with_runs(vec![
        TextRun::new("⏮ First").size(14).color_rgb(255, 255, 255),
    ]);
    writer.add_rich_textbox(slide5, 195, 258, 110, 25, vec![first_text.center()])?;
    writer.set_last_shape_hyperlink(slide5, first_id)?;

    let next_id = writer.add_hyperlink(Hyperlink::next_slide());
    writer.add_styled_shape(slide5, ShapeType::Rectangle, 330, 250, 120, 40, nav_style)?;
    let next_text = Paragraph::with_runs(vec![
        TextRun::new("Next ▶").size(14).color_rgb(255, 255, 255),
    ]);
    writer.add_rich_textbox(slide5, 335, 258, 110, 25, vec![next_text.center()])?;
    writer.set_last_shape_hyperlink(slide5, next_id)?;

    writer.set_slide_notes(
        slide5,
        "Interactive elements:\n\
         - Web links open in browser\n\
         - Navigation buttons move between slides\n\
         - Test in Slideshow mode",
    )?;

    // === SLIDE 6: Summary ===
    let slide6 = writer.add_slide()?;

    let summary_title = Paragraph::with_runs(vec![
        TextRun::new("Feature Summary")
            .bold()
            .size(36)
            .color_rgb(0, 51, 102),
    ]);
    writer.add_rich_textbox(slide6, 50, 50, 600, 60, vec![summary_title.center()])?;

    let features = vec![
        "✓ Text Formatting (bold, italic, underline, colors, sizes)",
        "✓ Shape Types (rectangles, ellipses, lines, arrows)",
        "✓ Shape Styling (fills, borders, shadows, gradients)",
        "✓ Hyperlinks (URLs, file links, slide navigation)",
        "✓ Speaker Notes (simple and rich text)",
        "✓ Custom Fonts",
    ];

    let mut y = 130;
    for feature in features {
        let para = Paragraph::with_runs(vec![TextRun::new(feature).size(18)]);
        writer.add_rich_textbox(slide6, 100, y, 500, 35, vec![para])?;
        y += 45;
    }

    writer.set_slide_notes(
        slide6,
        "FINAL SLIDE NOTES:\n\n\
         This presentation demonstrated all major features of the PPT writer:\n\n\
         1. Rich text formatting with multiple styles per paragraph\n\
         2. Various shape types with positioning\n\
         3. Shape styling including fills, borders, shadows, and gradients\n\
         4. Working hyperlinks for web and navigation\n\
         5. Speaker notes for each slide\n\n\
         All features follow MS-PPT and MS-ODRAW specifications.",
    )?;

    println!("  Slides: {}", writer.slide_count());
    println!("  Hyperlinks: {}", writer.hyperlink_count());
    println!("  Fonts: {}", writer.font_count());

    writer.save("output/08_comprehensive.ppt")?;
    println!("  ✓ output/08_comprehensive.ppt");
    Ok(())
}
