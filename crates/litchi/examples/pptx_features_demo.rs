//! Demonstration of PPTX features including Tables, Group Shapes, Comments, Sections, and Media.
//!
//! Run this example with:
//! ```bash
//! cargo run --example pptx_features_demo
//! ```
//!
//! This will generate `pptx_features_demo.pptx` in the project root directory.
//! Open it in Microsoft PowerPoint to verify all features work correctly.

use litchi::ooxml::pptx::{Package, Section, SectionList};
use std::fs;
use std::path::Path;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("Creating PPTX features demonstration...\n");

    // Create a new presentation package
    let mut pkg = Package::new()?;

    // Get mutable presentation for editing
    let pres = pkg.presentation_mut()?;

    // Set presentation metadata
    pres.set_slide_size(9144000, 6858000); // Standard 4:3 (10" x 7.5" in EMUs)

    // ========================================================================
    // Slide 1: Title Slide
    // ========================================================================
    println!("Creating Slide 1: Title Slide");
    let slide1 = pres.add_slide()?;
    slide1.set_title("PPTX Features Demo");
    slide1.add_text_box(
        "Demonstrating Tables, Groups, Comments, Sections, and Media",
        914400,  // x: 1 inch from left
        3429000, // y: 3.75 inches from top
        7315200, // width: 8 inches
        914400,  // height: 1 inch
    );

    // Add a comment to the title slide
    slide1.add_comment(
        0,
        "This is a demo presentation showcasing new PPTX features!",
        914400,
        914400,
    );

    // ========================================================================
    // Slide 2: Table Demonstration
    // ========================================================================
    println!("Creating Slide 2: Table Demonstration");
    let slide2 = pres.add_slide()?;
    slide2.set_title("Table Feature");

    // Create a simple data table
    let table_data = vec![
        vec![
            "Feature".to_string(),
            "Status".to_string(),
            "Notes".to_string(),
        ],
        vec![
            "Tables".to_string(),
            "✓ Implemented".to_string(),
            "Full read/write".to_string(),
        ],
        vec![
            "Group Shapes".to_string(),
            "✓ Implemented".to_string(),
            "Nested groups".to_string(),
        ],
        vec![
            "Comments".to_string(),
            "✓ Implemented".to_string(),
            "Author support".to_string(),
        ],
        vec![
            "Sections".to_string(),
            "✓ Implemented".to_string(),
            "Slide organization".to_string(),
        ],
        vec![
            "Audio/Video".to_string(),
            "✓ Implemented".to_string(),
            "Multiple formats".to_string(),
        ],
    ];

    slide2.add_table(
        table_data, 914400,  // x: 1 inch
        1828800, // y: 2 inches
        7315200, // width: 8 inches
        2743200, // height: 3 inches
    );

    // Add a comment explaining the table
    slide2.add_comment(
        0,
        "This table shows the implementation status of new features.",
        914400,
        1828800,
    );

    // ========================================================================
    // Slide 3: Group Shapes Demonstration
    // ========================================================================
    println!("Creating Slide 3: Group Shapes Demonstration");
    let slide3 = pres.add_slide()?;
    slide3.set_title("Group Shapes Feature");

    // Create a group with multiple shapes
    let group_idx = slide3.add_group(
        1828800, // x: 2 inches
        1828800, // y: 2 inches
        5486400, // width: 6 inches
        3657600, // height: 4 inches
    );

    // Add shapes to the group
    // Red rectangle (top-left)
    slide3.add_rectangle_to_group(
        group_idx,
        1828800,                    // x position within group
        1828800,                    // y position within group
        1828800,                    // width: 2 inches
        1371600,                    // height: 1.5 inches
        Some("FF6B6B".to_string()), // Coral red
    );

    // Green rectangle (top-right)
    slide3.add_rectangle_to_group(
        group_idx,
        4114800,                    // x position
        1828800,                    // y position
        1828800,                    // width
        1371600,                    // height
        Some("4ECDC4".to_string()), // Teal green
    );

    // Blue ellipse (bottom-center)
    slide3.add_ellipse_to_group(
        group_idx,
        2971800,                    // x position (centered)
        3657600,                    // y position
        2286000,                    // width: 2.5 inches
        1371600,                    // height: 1.5 inches
        Some("45B7D1".to_string()), // Sky blue
    );

    // Add a text box label to the group
    slide3.add_text_box_to_group(
        group_idx,
        "Grouped Shapes",
        2514600, // x
        5486400, // y
        3200400, // width
        457200,  // height
    );

    // Add explanatory text below
    slide3.add_text_box(
        "The shapes above are grouped together and can be manipulated as a single unit.",
        914400,
        5943600,
        7315200,
        457200,
    );

    // ========================================================================
    // Slide 4: Audio Demonstration
    // ========================================================================
    println!("Creating Slide 4: Audio Demonstration");
    let slide4 = pres.add_slide()?;
    slide4.set_title("Audio Feature");

    slide4.add_text_box(
        "This slide contains embedded audio files.\nClick the audio icons to play.",
        914400,
        1828800,
        7315200,
        914400,
    );

    // Load and add MP3 audio
    let mp3_path = Path::new("file_example_MP3_700KB.mp3");
    if mp3_path.exists() {
        let mp3_data = fs::read(mp3_path)?;
        slide4.add_audio(
            mp3_data, 1371600, // x: 1.5 inches
            3200400, // y: 3.5 inches
            914400,  // width: 1 inch
            914400,  // height: 1 inch
        );
        slide4.add_text_box("MP3 Audio", 1143000, 4200400, 1371600, 457200);
        println!("  - Added MP3 audio");
    } else {
        println!("  - Warning: MP3 file not found, skipping");
    }

    // Load and add WAV audio
    let wav_path = Path::new("file_example_WAV_1MG.wav");
    if wav_path.exists() {
        let wav_data = fs::read(wav_path)?;
        slide4.add_audio(
            wav_data, 3886200, // x: 4.25 inches
            3200400, // y: 3.5 inches
            914400,  // width
            914400,  // height
        );
        slide4.add_text_box("WAV Audio", 3657600, 4200400, 1371600, 457200);
        println!("  - Added WAV audio");
    } else {
        println!("  - Warning: WAV file not found, skipping");
    }

    // ========================================================================
    // Slide 5: Video Demonstration
    // ========================================================================
    println!("Creating Slide 5: Video Demonstration");
    let slide5 = pres.add_slide()?;
    slide5.set_title("Video Feature");

    slide5.add_text_box(
        "This slide contains an embedded video.\nClick to play the video.",
        914400,
        1600200,
        7315200,
        457200,
    );

    // Load and add MP4 video
    let video_path = Path::new("ForBiggerMeltdowns.mp4");
    if video_path.exists() {
        let video_data = fs::read(video_path)?;
        slide5.add_video(
            video_data, 1828800, // x: 2 inches
            2286000, // y: 2.5 inches
            5486400, // width: 6 inches
            3086100, // height: ~3.4 inches (16:9 aspect)
        );
        println!("  - Added MP4 video");
    } else {
        println!("  - Warning: Video file not found, skipping");
    }

    // ========================================================================
    // Slide 6: Comments Demonstration
    // ========================================================================
    println!("Creating Slide 6: Comments Demonstration");
    let slide6 = pres.add_slide()?;
    slide6.set_title("Comments Feature");

    slide6.add_text_box(
        "This slide has multiple comments attached.\nOpen the Comments pane in PowerPoint to view them.",
        914400,
        1828800,
        7315200,
        914400,
    );

    // Add multiple comments at different positions
    slide6.add_comment(
        0,
        "Comment 1: This is the first comment on this slide.",
        914400,
        3200400,
    );
    slide6.add_comment(
        0,
        "Comment 2: Comments can be placed at specific positions.",
        4572000,
        3200400,
    );
    slide6.add_comment(
        0,
        "Comment 3: Each comment is associated with an author.",
        914400,
        4572000,
    );

    // Add a visual marker where comments are
    slide6.add_rectangle(914400, 3200400, 228600, 228600, Some("FFD93D".to_string())); // Yellow marker
    slide6.add_rectangle(4572000, 3200400, 228600, 228600, Some("FFD93D".to_string()));
    slide6.add_rectangle(914400, 4572000, 228600, 228600, Some("FFD93D".to_string()));

    slide6.add_text_box(
        "Yellow markers indicate comment positions",
        914400,
        5486400,
        7315200,
        457200,
    );

    // ========================================================================
    // Slide 7: Summary with Table
    // ========================================================================
    println!("Creating Slide 7: Summary");
    let slide7 = pres.add_slide()?;
    slide7.set_title("Summary");

    let summary_data = vec![
        vec!["Slide".to_string(), "Feature Demonstrated".to_string()],
        vec!["1".to_string(), "Title + Comments".to_string()],
        vec!["2".to_string(), "Tables".to_string()],
        vec!["3".to_string(), "Group Shapes".to_string()],
        vec!["4".to_string(), "Audio (MP3, WAV)".to_string()],
        vec!["5".to_string(), "Video (MP4)".to_string()],
        vec!["6".to_string(), "Multiple Comments".to_string()],
        vec!["7".to_string(), "Summary Table".to_string()],
    ];

    slide7.add_table_with_options(
        summary_data,
        1828800, // x: 2 inches
        1828800, // y: 2 inches
        5486400, // width: 6 inches
        3200400, // height: 3.5 inches
        None,    // auto column widths
        None,    // auto row heights
        true,    // first row as header
        true,    // banded rows
    );

    // ========================================================================
    // Create Sections (Note: Sections are metadata, created separately)
    // ========================================================================
    println!("\nCreating sections...");
    let mut sections = SectionList::new();

    // Introduction section (slides 1)
    sections.add_section(
        Section::new("Introduction", "1").with_slides(vec![256]), // First slide ID
    );

    // Features section (slides 2-6)
    sections.add_section(
        Section::new("Feature Demonstrations", "2").with_slides(vec![257, 258, 259, 260, 261]),
    );

    // Summary section (slide 7)
    sections.add_section(Section::new("Summary", "3").with_slides(vec![262]));

    // Generate section XML (for reference - would be included in presentation.xml)
    let _section_xml = sections.to_xml()?;
    println!("  - Created {} sections", sections.len());

    // ========================================================================
    // Save the presentation
    // ========================================================================
    let output_path = "pptx_features_demo.pptx";
    println!("\nSaving presentation to {}...", output_path);

    // Save the package (mutable reference goes out of scope)
    pkg.save(output_path)?;

    println!("\n✓ Presentation created successfully!");
    println!(
        "\nOpen '{}' in Microsoft PowerPoint to verify:",
        output_path
    );
    println!("  - Slide 1: Title slide with a comment");
    println!("  - Slide 2: Table with feature status");
    println!("  - Slide 3: Group shapes (colored rectangles and ellipse)");
    println!("  - Slide 4: Audio files (MP3 and WAV)");
    println!("  - Slide 5: Video file (MP4)");
    println!("  - Slide 6: Multiple comments with position markers");
    println!("  - Slide 7: Summary table");

    Ok(())
}
