#![allow(
    clippy::print_stdout,
    reason = "this command-line example intentionally prints its results"
)]

//! Comprehensive PPT shapes showcase
//!
//! This example demonstrates creating a `PowerPoint` presentation with various
//! `Escher` shapes including rectangles, ellipses, lines, text boxes, arrows,
//! stars, callouts, flowchart shapes, and more - all with proper styling.
//!
//! Run with: `cargo run --example ppt_shapes_showcase`

use litchi_ppt::writer::{
    FillStyle, LineStyleConfig, ShadowStyle, ShapeColor, ShapeStyle, ShapeType, Writer,
};
use std::error::Error;

fn main() -> Result<(), Box<dyn Error>> {
    println!("Creating PPT file with comprehensive shape examples...");

    let mut ppt = Writer::new();

    // Slide 1: Basic shapes with solid fills
    create_basic_shapes_slide(&mut ppt)?;

    // Slide 2: Block arrows with different colors
    create_arrows_slide(&mut ppt)?;

    // Slide 3: Stars and banners
    create_stars_slide(&mut ppt)?;

    // Slide 4: Flowchart shapes
    create_flowchart_slide(&mut ppt)?;

    // Slide 5: Special shapes (hearts, diamonds, triangles)
    create_special_shapes_slide(&mut ppt)?;

    // Slide 6: Lines with various styles
    create_lines_slide(&mut ppt)?;

    // Slide 7: Shapes with thick borders & custom colors
    create_styled_shapes_slide(&mut ppt)?;

    // Slide 8: Gradient fills
    create_gradient_slide(&mut ppt)?;

    // Slide 9: Shadow effects
    create_shadow_slide(&mut ppt)?;

    // Slide 10: Transparency and opacity
    create_transparency_slide(&mut ppt)?;

    // Save the presentation
    let output_path = "output/ppt_shapes_showcase.ppt";
    ppt.save(output_path)?;

    println!("✓ Created: {output_path}");
    println!("  10 slides with 80+ shapes");
    println!(
        "  Includes: basic shapes, arrows, stars, flowchart, gradients, shadows, transparency"
    );
    println!("  Advanced styling: solid fills, gradients, drop shadows, custom shadows, opacity");
    println!("  Open in Microsoft PowerPoint to verify shapes render correctly");

    Ok(())
}

fn create_basic_shapes_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Basic Shapes with Solid Fills")?;

    // Blue rectangle
    let blue_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::BLUE))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 100, 150, 80, blue_style)?;
    ppt.add_textbox(slide, 50, 190, 150, 25, "Blue Rectangle")?;

    // Red ellipse
    let red_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::RED))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 230, 100, 150, 80, red_style)?;
    ppt.add_textbox(slide, 230, 190, 150, 25, "Red Ellipse")?;

    // Green rounded rectangle
    let green_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::GREEN))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        410,
        100,
        150,
        80,
        green_style,
    )?;
    ppt.add_textbox(slide, 410, 190, 150, 25, "Rounded Rect")?;

    // Yellow diamond
    let yellow_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::YELLOW))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 50, 250, 120, 120, yellow_style)?;
    ppt.add_textbox(slide, 50, 380, 120, 25, "Diamond")?;

    // Orange triangle
    let orange_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::ORANGE))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Triangle, 200, 250, 120, 120, orange_style)?;
    ppt.add_textbox(slide, 200, 380, 120, 25, "Triangle")?;

    // Purple hexagon (custom color)
    let purple = ShapeColor::rgb(128, 0, 128);
    let purple_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(purple))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 350, 250, 120, 120, purple_style)?;
    ppt.add_textbox(slide, 350, 380, 120, 25, "Purple")?;

    // Pink circle (custom color)
    let pink = ShapeColor::rgb(255, 192, 203);
    let pink_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(pink))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 500, 250, 120, 120, pink_style)?;
    ppt.add_textbox(slide, 500, 380, 120, 25, "Pink Circle")?;

    Ok(())
}

fn create_arrows_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Block Arrows with Colors")?;

    // Right arrow - Red
    let red_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::RED))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(139, 0, 0),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Arrow, 50, 100, 180, 80, red_style)?;
    ppt.add_textbox(slide, 50, 190, 180, 25, "Right Arrow")?;

    // Heart - Pink
    let pink = ShapeColor::rgb(255, 105, 180);
    let heart_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(pink))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(199, 21, 133),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Heart, 270, 100, 100, 90, heart_style)?;
    ppt.add_textbox(slide, 270, 200, 100, 25, "Heart")?;

    // Star - Yellow
    let star_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::YELLOW))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::ORANGE,
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Star, 410, 100, 110, 110, star_style)?;
    ppt.add_textbox(slide, 410, 220, 110, 25, "Star")?;

    // Custom teal arrow
    let teal = ShapeColor::rgb(0, 128, 128);
    let teal_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(teal))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 100, 100),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Arrow, 50, 270, 180, 80, teal_style)?;
    ppt.add_textbox(slide, 50, 360, 180, 25, "Teal Arrow")?;

    // Custom lavender ellipse
    let lavender = ShapeColor::rgb(230, 230, 250);
    let lavender_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(lavender))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(147, 112, 219),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 270, 270, 120, 90, lavender_style)?;
    ppt.add_textbox(slide, 270, 370, 120, 25, "Lavender")?;

    // Custom cyan star
    let cyan = ShapeColor::rgb(0, 255, 255);
    let cyan_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(cyan))
        .with_line(LineStyleConfig::with_color_and_width(ShapeColor::BLUE, 2.5));
    ppt.add_styled_shape(slide, ShapeType::Star, 430, 270, 110, 110, cyan_style)?;
    ppt.add_textbox(slide, 430, 390, 110, 25, "Cyan Star")?;

    Ok(())
}

fn create_stars_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Stars, Hearts, and Special Shapes")?;

    // Gold star
    let gold = ShapeColor::rgb(255, 215, 0);
    let gold_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(gold))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::ORANGE,
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Star, 50, 100, 110, 110, gold_style)?;
    ppt.add_textbox(slide, 50, 220, 110, 25, "Gold Star")?;

    // Silver star
    let silver = ShapeColor::rgb(192, 192, 192);
    let silver_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(silver))
        .with_line(LineStyleConfig::with_color_and_width(ShapeColor::GRAY, 2.0));
    ppt.add_styled_shape(slide, ShapeType::Star, 190, 100, 110, 110, silver_style)?;
    ppt.add_textbox(slide, 190, 220, 110, 25, "Silver Star")?;

    // Red heart
    let heart_red = ShapeColor::rgb(220, 20, 60);
    let red_heart_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(heart_red))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(139, 0, 0),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Heart, 330, 100, 100, 100, red_heart_style)?;
    ppt.add_textbox(slide, 330, 210, 100, 25, "Red Heart")?;

    // Purple heart
    let purple_heart = ShapeColor::rgb(138, 43, 226);
    let purple_heart_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(purple_heart))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(75, 0, 130),
            2.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::Heart,
        460,
        100,
        100,
        100,
        purple_heart_style,
    )?;
    ppt.add_textbox(slide, 460, 210, 100, 25, "Purple Heart")?;

    // Lime green diamond
    let lime = ShapeColor::rgb(50, 205, 50);
    let lime_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(lime))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(34, 139, 34),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 100, 280, 110, 110, lime_style)?;
    ppt.add_textbox(slide, 100, 400, 110, 25, "Lime Diamond")?;

    // Sky blue rounded rect
    let sky_blue = ShapeColor::rgb(135, 206, 235);
    let sky_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(sky_blue))
        .with_line(LineStyleConfig::with_color_and_width(ShapeColor::BLUE, 2.0));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        250,
        280,
        140,
        90,
        sky_style,
    )?;
    ppt.add_textbox(slide, 250, 380, 140, 25, "Sky Blue")?;

    // Coral triangle
    let coral = ShapeColor::rgb(255, 127, 80);
    let coral_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(coral))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(255, 99, 71),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Triangle, 430, 280, 110, 110, coral_style)?;
    ppt.add_textbox(slide, 430, 400, 110, 25, "Coral Triangle")?;

    Ok(())
}

fn create_flowchart_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Flowchart Shapes")?;

    // Note: Using basic shapes as flowchart equivalents since ShapeType doesn't expose flowchart shapes
    // In real implementation, we'd use shape_type::FLOWCHART_* constants

    // Start/End (rounded rectangle) - Light green
    let light_green = ShapeColor::rgb(144, 238, 144);
    let start_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(light_green))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 128, 0),
            2.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        300,
        80,
        140,
        60,
        start_style,
    )?;
    ppt.add_textbox(slide, 320, 100, 100, 20, "Start")?;

    // Connecting line
    ppt.add_line(slide, 370, 140, 370, 180)?;

    // Process (rectangle) - Light blue
    let light_blue = ShapeColor::rgb(173, 216, 230);
    let process_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(light_blue))
        .with_line(LineStyleConfig::with_color_and_width(ShapeColor::BLUE, 2.0));
    ppt.add_styled_shape(
        slide,
        ShapeType::Rectangle,
        300,
        180,
        140,
        60,
        process_style,
    )?;
    ppt.add_textbox(slide, 320, 200, 100, 20, "Process")?;

    // Connecting line
    ppt.add_line(slide, 370, 240, 370, 280)?;

    // Decision (diamond) - Light yellow
    let light_yellow = ShapeColor::rgb(255, 255, 224);
    let decision_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(light_yellow))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::ORANGE,
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 310, 280, 120, 80, decision_style)?;
    ppt.add_textbox(slide, 330, 310, 80, 20, "Decision")?;

    // Connecting lines
    ppt.add_line(slide, 370, 360, 370, 400)?;

    // End (rounded rectangle) - Light coral
    let light_coral = ShapeColor::rgb(240, 128, 128);
    let end_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(light_coral))
        .with_line(LineStyleConfig::with_color_and_width(ShapeColor::RED, 2.0));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        300,
        400,
        140,
        60,
        end_style,
    )?;
    ppt.add_textbox(slide, 330, 420, 80, 20, "End")?;

    // Side shapes for variety
    // Data (parallelogram approximation - using diamond)
    let peach = ShapeColor::rgb(255, 218, 185);
    let data_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(peach))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::ORANGE,
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 520, 180, 100, 70, data_style)?;
    ppt.add_textbox(slide, 540, 205, 60, 20, "Data")?;

    Ok(())
}

fn create_special_shapes_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "More Special Shapes")?;

    // Mint green rectangle
    let mint = ShapeColor::rgb(152, 251, 152);
    let mint_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(mint))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 128, 0),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 100, 140, 90, mint_style)?;
    ppt.add_textbox(slide, 70, 200, 100, 20, "Mint")?;

    // Rose ellipse
    let rose = ShapeColor::rgb(255, 182, 193);
    let rose_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(rose))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(219, 112, 147),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 220, 100, 120, 90, rose_style)?;
    ppt.add_textbox(slide, 240, 200, 80, 20, "Rose")?;

    // Turquoise diamond
    let turquoise = ShapeColor::rgb(64, 224, 208);
    let turquoise_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(turquoise))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 206, 209),
            2.5,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::Diamond,
        370,
        100,
        100,
        100,
        turquoise_style,
    )?;
    ppt.add_textbox(slide, 385, 210, 70, 20, "Turquoise")?;

    // Plum rounded rect
    let plum = ShapeColor::rgb(221, 160, 221);
    let plum_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(plum))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(186, 85, 211),
            2.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        500,
        100,
        120,
        90,
        plum_style,
    )?;
    ppt.add_textbox(slide, 525, 200, 70, 20, "Plum")?;

    // Khaki triangle
    let khaki = ShapeColor::rgb(240, 230, 140);
    let khaki_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(khaki))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(189, 183, 107),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Triangle, 80, 270, 100, 100, khaki_style)?;
    ppt.add_textbox(slide, 95, 380, 70, 20, "Khaki")?;

    // Indigo star
    let indigo = ShapeColor::rgb(75, 0, 130);
    let indigo_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(indigo))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(138, 43, 226),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Star, 220, 270, 100, 100, indigo_style)?;
    ppt.add_textbox(slide, 230, 380, 80, 20, "Indigo Star")?;

    // Salmon heart
    let salmon = ShapeColor::rgb(250, 128, 114);
    let salmon_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(salmon))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(233, 150, 122),
            2.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Heart, 350, 270, 90, 90, salmon_style)?;
    ppt.add_textbox(slide, 355, 370, 80, 20, "Salmon")?;

    // Olive arrow
    let olive = ShapeColor::rgb(128, 128, 0);
    let olive_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(olive))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(85, 107, 47),
            2.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Arrow, 470, 280, 150, 70, olive_style)?;
    ppt.add_textbox(slide, 500, 360, 90, 20, "Olive Arrow")?;

    Ok(())
}
fn create_lines_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Lines with Various Styles")?;

    // Note: Lines are drawn with Line shape type
    // Black thick line
    ppt.add_line(slide, 50, 100, 250, 100)?;
    ppt.add_textbox(slide, 100, 110, 100, 20, "Thick Line")?;

    // Diagonal line
    ppt.add_line(slide, 300, 100, 450, 180)?;
    ppt.add_textbox(slide, 350, 150, 80, 20, "Diagonal")?;

    // Vertical line
    ppt.add_line(slide, 500, 100, 500, 200)?;
    ppt.add_textbox(slide, 510, 140, 70, 20, "Vertical")?;

    // Cross lines
    ppt.add_line(slide, 100, 250, 200, 350)?;
    ppt.add_line(slide, 200, 250, 100, 350)?;
    ppt.add_textbox(slide, 120, 360, 60, 20, "Cross")?;

    // Box made of lines
    ppt.add_line(slide, 300, 250, 450, 250)?; // Top
    ppt.add_line(slide, 450, 250, 450, 350)?; // Right
    ppt.add_line(slide, 450, 350, 300, 350)?; // Bottom
    ppt.add_line(slide, 300, 350, 300, 250)?; // Left
    ppt.add_textbox(slide, 340, 390, 70, 20, "Box Lines")?;

    // Star pattern with lines
    let cx = 550;
    let cy = 300;
    ppt.add_line(slide, cx, cy - 50, cx, cy + 50)?; // Vertical
    ppt.add_line(slide, cx - 50, cy, cx + 50, cy)?; // Horizontal
    ppt.add_line(slide, cx - 35, cy - 35, cx + 35, cy + 35)?; // Diagonal 1
    ppt.add_line(slide, cx + 35, cy - 35, cx - 35, cy + 35)?; // Diagonal 2
    ppt.add_textbox(slide, 510, 360, 80, 20, "Star Pattern")?;

    Ok(())
}

fn create_styled_shapes_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(
        slide,
        50,
        20,
        650,
        40,
        "Shapes with Thick Borders & Custom Colors",
    )?;

    // Crimson rectangle with thick border
    let crimson = ShapeColor::rgb(220, 20, 60);
    let crimson_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(crimson))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(139, 0, 0),
            4.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::Rectangle,
        50,
        100,
        150,
        100,
        crimson_style,
    )?;
    ppt.add_textbox(slide, 60, 210, 130, 20, "Thick Border")?;

    // Navy ellipse
    let navy = ShapeColor::rgb(0, 0, 128);
    let navy_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(navy))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(25, 25, 112),
            3.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 230, 100, 150, 100, navy_style)?;
    ppt.add_textbox(slide, 260, 210, 90, 20, "Navy Blue")?;

    // Forest green diamond
    let forest = ShapeColor::rgb(34, 139, 34);
    let forest_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(forest))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(0, 100, 0),
            4.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 410, 100, 120, 120, forest_style)?;
    ppt.add_textbox(slide, 425, 230, 90, 20, "Forest Green")?;

    // Maroon star
    let maroon = ShapeColor::rgb(128, 0, 0);
    let maroon_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(maroon))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(165, 42, 42),
            3.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Star, 80, 280, 100, 100, maroon_style)?;
    ppt.add_textbox(slide, 90, 390, 80, 20, "Maroon")?;

    // Chocolate brown rounded rect
    let chocolate = ShapeColor::rgb(210, 105, 30);
    let chocolate_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(chocolate))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(139, 69, 19),
            3.0,
        ));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        220,
        280,
        140,
        90,
        chocolate_style,
    )?;
    ppt.add_textbox(slide, 240, 380, 100, 20, "Chocolate")?;

    // Steel blue triangle
    let steel = ShapeColor::rgb(70, 130, 180);
    let steel_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(steel))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(100, 149, 237),
            3.5,
        ));
    ppt.add_styled_shape(slide, ShapeType::Triangle, 390, 280, 110, 110, steel_style)?;
    ppt.add_textbox(slide, 405, 400, 80, 20, "Steel Blue")?;

    // Dark orange heart
    let dark_orange = ShapeColor::rgb(255, 140, 0);
    let dark_orange_style = ShapeStyle::default()
        .with_fill(FillStyle::solid(dark_orange))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(255, 69, 0),
            3.0,
        ));
    ppt.add_styled_shape(slide, ShapeType::Heart, 530, 290, 90, 90, dark_orange_style)?;
    ppt.add_textbox(slide, 530, 390, 90, 20, "Dark Orange")?;

    Ok(())
}
// Advanced slide functions for gradients, shadows, and transparency

fn create_gradient_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Gradient Fills - Various Angles")?;

    // Horizontal gradient (0 degrees) - Blue to Cyan
    let grad1 = FillStyle::gradient(ShapeColor::BLUE, ShapeColor::rgb(0, 255, 255), 0);
    let grad1_style =
        ShapeStyle::default()
            .with_fill(grad1)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 100, 150, 80, grad1_style)?;
    ppt.add_textbox(slide, 50, 190, 150, 25, "0° Horizontal")?;

    // Vertical gradient (90 degrees) - Red to Yellow
    let grad2 = FillStyle::gradient(ShapeColor::RED, ShapeColor::YELLOW, 90);
    let grad2_style =
        ShapeStyle::default()
            .with_fill(grad2)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 230, 100, 150, 80, grad2_style)?;
    ppt.add_textbox(slide, 230, 190, 150, 25, "90° Vertical")?;

    // Diagonal gradient (45 degrees) - Green to Blue
    let grad3 = FillStyle::gradient(ShapeColor::GREEN, ShapeColor::BLUE, 45);
    let grad3_style =
        ShapeStyle::default()
            .with_fill(grad3)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 410, 100, 150, 80, grad3_style)?;
    ppt.add_textbox(slide, 410, 190, 150, 25, "45° Diagonal")?;

    // Purple to Pink gradient ellipse
    let purple = ShapeColor::rgb(128, 0, 128);
    let pink = ShapeColor::rgb(255, 192, 203);
    let grad4 = FillStyle::gradient(purple, pink, 135);
    let grad4_style =
        ShapeStyle::default()
            .with_fill(grad4)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 80, 250, 120, 120, grad4_style)?;
    ppt.add_textbox(slide, 85, 380, 110, 25, "135° Ellipse")?;

    // Orange to Red gradient star
    let grad5 = FillStyle::gradient(ShapeColor::ORANGE, ShapeColor::RED, 180);
    let grad5_style =
        ShapeStyle::default()
            .with_fill(grad5)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Star, 240, 250, 110, 110, grad5_style)?;
    ppt.add_textbox(slide, 250, 370, 90, 25, "180° Star")?;

    // Teal to Lime gradient diamond
    let teal = ShapeColor::rgb(0, 128, 128);
    let lime = ShapeColor::rgb(50, 205, 50);
    let grad6 = FillStyle::gradient(teal, lime, 270);
    let grad6_style =
        ShapeStyle::default()
            .with_fill(grad6)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Diamond, 390, 250, 110, 110, grad6_style)?;
    ppt.add_textbox(slide, 395, 370, 100, 25, "270° Diamond")?;

    // Custom gradient - Navy to Sky Blue
    let navy = ShapeColor::rgb(0, 0, 128);
    let sky_blue = ShapeColor::rgb(135, 206, 235);
    let grad7 = FillStyle::gradient(navy, sky_blue, 315);
    let grad7_style =
        ShapeStyle::default()
            .with_fill(grad7)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(
        slide,
        ShapeType::RoundRectangle,
        530,
        260,
        120,
        90,
        grad7_style,
    )?;
    ppt.add_textbox(slide, 540, 360, 100, 25, "315° Round")?;

    Ok(())
}

fn create_shadow_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(
        slide,
        50,
        20,
        650,
        40,
        "Shadow Effects - Drop Shadows & Custom",
    )?;

    // Default drop shadow - Blue rectangle
    let shadow1 = ShadowStyle::drop_shadow();
    let style1 = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::BLUE))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow1);
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 100, 140, 90, style1)?;
    ppt.add_textbox(slide, 60, 200, 120, 25, "Drop Shadow")?;

    // Custom shadow - offset right/down - Red ellipse
    let shadow2 = ShadowStyle::custom(ShapeColor::rgb(150, 0, 0), 38100, 38100, 60); // 3pt offset
    let style2 = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::RED))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow2);
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 220, 100, 130, 90, style2)?;
    ppt.add_textbox(slide, 235, 200, 100, 25, "Red Shadow")?;

    // Large shadow - Green star
    let shadow3 = ShadowStyle::custom(ShapeColor::rgb(0, 100, 0), 50800, 50800, 70); // 4pt offset
    let style3 = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::GREEN))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow3);
    ppt.add_styled_shape(slide, ShapeType::Star, 390, 100, 110, 110, style3)?;
    ppt.add_textbox(slide, 400, 220, 90, 25, "Large Shadow")?;

    // Light shadow - Yellow diamond
    let shadow4 = ShadowStyle::custom(ShapeColor::rgb(200, 200, 0), 25400, 25400, 30); // 2pt, light
    let style4 = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::YELLOW))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow4);
    ppt.add_styled_shape(slide, ShapeType::Diamond, 540, 100, 100, 100, style4)?;
    ppt.add_textbox(slide, 545, 210, 90, 25, "Light Shadow")?;

    // Purple shadow on orange shape
    let shadow5 = ShadowStyle::custom(ShapeColor::rgb(128, 0, 128), 38100, 38100, 80);
    let style5 = ShapeStyle::default()
        .with_fill(FillStyle::solid(ShapeColor::ORANGE))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow5);
    ppt.add_styled_shape(slide, ShapeType::Triangle, 90, 280, 110, 110, style5)?;
    ppt.add_textbox(slide, 95, 400, 100, 25, "Purple Shadow")?;

    // Black shadow on cyan shape
    let shadow6 = ShadowStyle::custom(ShapeColor::BLACK, 38100, 38100, 50);
    let cyan = ShapeColor::rgb(0, 255, 255);
    let style6 = ShapeStyle::default()
        .with_fill(FillStyle::solid(cyan))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow6);
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 240, 280, 120, 90, style6)?;
    ppt.add_textbox(slide, 250, 380, 100, 25, "Black Shadow")?;

    // Gray shadow on pink heart
    let shadow7 = ShadowStyle::custom(ShapeColor::GRAY, 25400, 25400, 40);
    let pink = ShapeColor::rgb(255, 192, 203);
    let style7 = ShapeStyle::default()
        .with_fill(FillStyle::solid(pink))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::BLACK,
            1.5,
        ))
        .with_shadow(shadow7);
    ppt.add_styled_shape(slide, ShapeType::Heart, 400, 280, 90, 90, style7)?;
    ppt.add_textbox(slide, 405, 380, 80, 25, "Gray Shadow")?;

    // Soft shadow on rounded rect
    let shadow8 = ShadowStyle::custom(ShapeColor::rgb(100, 100, 100), 25400, 25400, 25);
    let lavender = ShapeColor::rgb(230, 230, 250);
    let style8 = ShapeStyle::default()
        .with_fill(FillStyle::solid(lavender))
        .with_line(LineStyleConfig::with_color_and_width(
            ShapeColor::rgb(147, 112, 219),
            1.5,
        ))
        .with_shadow(shadow8);
    ppt.add_styled_shape(slide, ShapeType::RoundRectangle, 530, 280, 120, 90, style8)?;
    ppt.add_textbox(slide, 545, 380, 90, 25, "Soft Shadow")?;

    Ok(())
}

fn create_transparency_slide(ppt: &mut Writer) -> Result<(), Box<dyn Error>> {
    let slide = ppt.add_slide()?;

    // Title
    ppt.add_textbox(slide, 50, 20, 650, 40, "Transparency & Opacity Levels")?;

    // 100% opacity (fully opaque) - Blue
    let fill1 = FillStyle::solid(ShapeColor::BLUE).with_opacity(100);
    let style1 =
        ShapeStyle::default()
            .with_fill(fill1)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 100, 120, 80, style1)?;
    ppt.add_textbox(slide, 60, 190, 100, 20, "100% Opaque")?;

    // 80% opacity - Red
    let fill2 = FillStyle::solid(ShapeColor::RED).with_opacity(80);
    let style2 =
        ShapeStyle::default()
            .with_fill(fill2)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 200, 100, 120, 80, style2)?;
    ppt.add_textbox(slide, 220, 190, 80, 20, "80% Opacity")?;

    // 60% opacity - Green
    let fill3 = FillStyle::solid(ShapeColor::GREEN).with_opacity(60);
    let style3 =
        ShapeStyle::default()
            .with_fill(fill3)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 350, 100, 120, 80, style3)?;
    ppt.add_textbox(slide, 370, 190, 80, 20, "60% Opacity")?;

    // 40% opacity - Yellow
    let fill4 = FillStyle::solid(ShapeColor::YELLOW).with_opacity(40);
    let style4 =
        ShapeStyle::default()
            .with_fill(fill4)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 500, 100, 120, 80, style4)?;
    ppt.add_textbox(slide, 520, 190, 80, 20, "40% Opacity")?;

    // 20% opacity - Orange
    let fill5 = FillStyle::solid(ShapeColor::ORANGE).with_opacity(20);
    let style5 =
        ShapeStyle::default()
            .with_fill(fill5)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Rectangle, 50, 240, 120, 80, style5)?;
    ppt.add_textbox(slide, 70, 330, 80, 20, "20% Opacity")?;

    // Overlapping circles to demonstrate transparency
    ppt.add_textbox(slide, 250, 240, 200, 20, "Overlapping Circles:")?;

    let purple = ShapeColor::rgb(128, 0, 128);
    let fill6 = FillStyle::solid(purple).with_opacity(50);
    let style6 =
        ShapeStyle::default()
            .with_fill(fill6)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::rgb(100, 0, 100),
                1.0,
            ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 280, 270, 100, 100, style6)?;

    let cyan = ShapeColor::rgb(0, 255, 255);
    let fill7 = FillStyle::solid(cyan).with_opacity(50);
    let style7 =
        ShapeStyle::default()
            .with_fill(fill7)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::rgb(0, 200, 200),
                1.0,
            ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 330, 270, 100, 100, style7)?;

    let yellow_transp = ShapeColor::YELLOW;
    let fill8 = FillStyle::solid(yellow_transp).with_opacity(50);
    let style8 =
        ShapeStyle::default()
            .with_fill(fill8)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::rgb(200, 200, 0),
                1.0,
            ));
    ppt.add_styled_shape(slide, ShapeType::Ellipse, 305, 310, 100, 100, style8)?;

    // Semi-transparent gradient
    let grad = FillStyle::gradient(ShapeColor::BLUE, ShapeColor::RED, 45).with_opacity(70);
    let style9 =
        ShapeStyle::default()
            .with_fill(grad)
            .with_line(LineStyleConfig::with_color_and_width(
                ShapeColor::BLACK,
                1.5,
            ));
    ppt.add_styled_shape(slide, ShapeType::Star, 500, 280, 110, 110, style9)?;
    ppt.add_textbox(slide, 505, 400, 100, 20, "Gradient 70%")?;

    Ok(())
}
