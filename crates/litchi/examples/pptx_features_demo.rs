//! PresentationML feature showcase using the current semantic writer and
//! typed package owners.
//!
//! Run with:
//! ```bash
//! cargo run --example pptx_features_demo
//! ```

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::comments::{
    Author as CommentAuthor, Comment, Comments, Conformance as CommentConformance,
    List as CommentList, store_presentation_comments,
};
use litchi_pptx::media_parts::{
    Conformance as MediaConformance, Kind as MediaKind, List as MediaList, Picture, Resource,
    Transform, store as store_slide_media,
};
use litchi_pptx::presentation::media::{Format as MediaFormat, Kind as AuthoringMediaKind};
use litchi_pptx::presentation_properties::metadata::sections::{List as SectionList, Section};
use litchi_pptx::{Error, MutablePresentation, MutableSlide, Package, Result};
use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::fs;
use std::path::Path;

const INCH: i64 = 914_400;
const SLIDE_WIDTH: i64 = 10 * INCH;
const SLIDE_HEIGHT: i64 = 7 * INCH + INCH / 2;

struct MediaAttachment {
    slide: usize,
    data: Vec<u8>,
    format: MediaFormat,
    name: String,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
}

fn main() -> std::result::Result<(), Box<dyn StdError>> {
    println!("Creating PPTX features demonstration...\n");

    let mut package = Package::new()?;
    let media = {
        let presentation = package.presentation_mut()?;
        presentation.set_slide_size(SLIDE_WIDTH, SLIDE_HEIGHT);
        build_slides(presentation)?
    };

    let sections = build_sections();
    println!(
        "Creating {} typed sections ({} XML bytes)...",
        sections.len(),
        sections.to_xml()?.len()
    );

    package.edit_opc(move |opc| {
        store_comments(opc)?;
        store_sections(opc, &sections.to_xml()?)?;
        store_media(opc, media)
    })?;

    let output_path = "pptx_features_demo.pptx";
    package.save(output_path)?;

    println!("\nPresentation created successfully: {output_path}");
    println!("  - semantic text and primitive-shape authoring");
    println!("  - table rows rendered as typed-writer row snapshots");
    println!("  - grouped-shape composition rendered as one visual scene");
    println!("  - typed legacy comments and optional audio/video parts");
    println!("  - typed PresentationML sections");
    Ok(())
}

fn build_slides(presentation: &mut MutablePresentation) -> Result<Vec<MediaAttachment>> {
    println!("Creating Slide 1: Title Slide");
    let slide = presentation.add_slide()?;
    slide.set_title("PPTX Features Demo");
    slide
        .add_text_box(
            "Demonstrating tables, grouped scenes, comments, sections, and media",
            INCH,
            3 * INCH + INCH / 4,
            8 * INCH,
            INCH,
        )
        .font_size(18.0)
        .bold(true);

    println!("Creating Slide 2: Table Demonstration");
    let slide = presentation.add_slide()?;
    slide.set_title("Table Feature");
    add_text_table(
        slide,
        &[
            &["Feature", "Status", "Notes"][..],
            &["Tables", "Implemented", "Row snapshots"][..],
            &["Group Shapes", "Implemented", "Visual composition"][..],
            &["Comments", "Implemented", "Typed package graph"][..],
            &["Sections", "Implemented", "Slide organization"][..],
            &["Audio/Video", "Implemented", "Media resources"][..],
        ],
        INCH,
        2 * INCH,
        8 * INCH,
        3 * INCH,
    );

    println!("Creating Slide 3: Group Shapes Demonstration");
    let slide = presentation.add_slide()?;
    slide.set_title("Grouped Scene Feature");
    // The current writer exposes semantic primitive shapes; compose the same
    // scene without the retired group-mutator facade.
    slide.add_rectangle(
        2 * INCH,
        2 * INCH,
        2 * INCH,
        3 * INCH / 2,
        Some("FF6B6B".into()),
    );
    slide.add_rectangle(
        5 * INCH,
        2 * INCH,
        2 * INCH,
        3 * INCH / 2,
        Some("4ECDC4".into()),
    );
    slide.add_ellipse(
        3 * INCH + INCH / 4,
        4 * INCH,
        2 * INCH + INCH / 2,
        3 * INCH / 2,
        Some("45B7D1".into()),
    );
    slide
        .add_text_box(
            "Grouped Shapes",
            2 * INCH + INCH / 2,
            5 * INCH + INCH / 2,
            5 * INCH,
            INCH / 2,
        )
        .bold(true)
        .font_size(16.0);
    slide.add_text_box(
        "The primitives form one grouped visual scene.",
        INCH,
        6 * INCH,
        8 * INCH,
        INCH / 2,
    );

    println!("Creating Slide 4: Audio Demonstration");
    let slide = presentation.add_slide()?;
    slide.set_title("Audio Feature");
    slide.add_text_box(
        "Optional audio resources are attached through the typed media package owner.",
        INCH,
        2 * INCH,
        8 * INCH,
        INCH,
    );
    let mut media = Vec::new();
    for (path, x, label) in [
        (
            Path::new("file_example_MP3_700KB.mp3"),
            3 * INCH / 2,
            "MP3 Audio",
        ),
        (
            Path::new("file_example_WAV_1MG.wav"),
            17 * INCH / 4,
            "WAV Audio",
        ),
    ] {
        if let Some(attachment) = read_media(path, 4, x, 7 * INCH / 2, label)? {
            slide.add_text_box(
                &format!("{} (typed media)", attachment.name),
                x - INCH / 4,
                5 * INCH,
                2 * INCH,
                INCH / 2,
            );
            println!("  - queued {}", attachment.name);
            media.push(attachment);
        } else {
            println!("  - {} not found, skipping", path.display());
        }
    }

    println!("Creating Slide 5: Video Demonstration");
    let slide = presentation.add_slide()?;
    slide.set_title("Video Feature");
    slide.add_text_box(
        "Optional video is attached through the typed media package owner.",
        INCH,
        2 * INCH,
        8 * INCH,
        INCH / 2,
    );
    let path = Path::new("ForBiggerMeltdowns.mp4");
    if let Some(attachment) = read_media(path, 5, 2 * INCH, 5 * INCH / 2, "MP4 Video")? {
        println!("  - queued {}", attachment.name);
        media.push(attachment);
    } else {
        println!("  - {} not found, skipping", path.display());
    }

    println!("Creating Slide 6: Comments Demonstration");
    let slide = presentation.add_slide()?;
    slide.set_title("Comments Feature");
    slide.add_text_box(
        "Typed comments are attached after the writer publishes the slide graph.",
        INCH,
        2 * INCH,
        8 * INCH,
        INCH,
    );
    for (x, y) in [
        (INCH, 7 * INCH / 2),
        (5 * INCH, 7 * INCH / 2),
        (INCH, 5 * INCH),
    ] {
        slide.add_rectangle(x, y, INCH / 4, INCH / 4, Some("FFD93D".into()));
    }
    slide.add_text_box(
        "Yellow markers indicate comment positions",
        INCH,
        6 * INCH,
        8 * INCH,
        INCH / 2,
    );

    println!("Creating Slide 7: Summary");
    let slide = presentation.add_slide()?;
    slide.set_title("Summary");
    add_text_table(
        slide,
        &[
            &["Slide", "Feature Demonstrated"][..],
            &["1", "Title + comments"][..],
            &["2", "Table row snapshots"][..],
            &["3", "Grouped visual scene"][..],
            &["4", "Audio resources"][..],
            &["5", "Video resource"][..],
            &["6", "Multiple comments"][..],
            &["7", "Summary table"][..],
        ],
        2 * INCH,
        2 * INCH,
        6 * INCH,
        7 * INCH / 2,
    );

    Ok(media)
}

fn add_text_table(
    slide: &mut MutableSlide,
    rows: &[&[&str]],
    x: i64,
    y: i64,
    width: i64,
    height: i64,
) {
    if rows.is_empty() {
        return;
    }
    let row_height = height / rows.len() as i64;
    for (index, row) in rows.iter().enumerate() {
        let row_y = y + index as i64 * row_height;
        slide.add_rectangle(
            x,
            row_y,
            width,
            row_height,
            Some(if index == 0 { "D9EAF7" } else { "F4F7FA" }.into()),
        );
        slide
            .add_text_box(
                &row.join("  |  "),
                x + INCH / 8,
                row_y + INCH / 8,
                width - INCH / 4,
                row_height - INCH / 8,
            )
            .font_size(if index == 0 { 13.0 } else { 11.0 })
            .bold(index == 0);
    }
}

fn read_media(
    path: &Path,
    slide: usize,
    x: i64,
    y: i64,
    _label: &str,
) -> Result<Option<MediaAttachment>> {
    if !path.exists() {
        return Ok(None);
    }
    let extension = path
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or_default();
    let format = MediaFormat::from_extension(extension);
    if format == MediaFormat::Unknown {
        return Err(Error::Invalid(format!(
            "unsupported media extension for {}",
            path.display()
        )));
    }
    let name = path
        .file_name()
        .and_then(|value| value.to_str())
        .unwrap_or("media")
        .to_owned();
    Ok(Some(MediaAttachment {
        slide,
        data: fs::read(path).map_err(|error| {
            Error::Invalid(format!("could not read {}: {error}", path.display()))
        })?,
        format,
        name,
        x,
        y,
        width: INCH,
        height: INCH,
    }))
}

fn build_sections() -> SectionList {
    let mut sections = SectionList::new();
    sections.add_section(Section::new("Introduction", "section-1").with_slides([256]));
    sections.add_section(
        Section::new("Feature Demonstrations", "section-2").with_slides([257, 258, 259, 260, 261]),
    );
    sections.add_section(Section::new("Summary", "section-3").with_slides([262]));
    sections
}

fn store_comments(opc: &mut OpcPackage) -> Result<()> {
    let mut author = CommentAuthor::new(1, "Litchi Demo", "LD");
    author.last_index = 5;
    let comments = Comments {
        author_relationship_id: "rIdCommentAuthors".into(),
        author_part_name: "/ppt/commentAuthors.xml".into(),
        authors: vec![author],
        slides: vec![
            CommentList {
                slide_part_name: "/ppt/slides/slide1.xml".into(),
                relationship_id: "rIdComments1".into(),
                part_name: "/ppt/comments/comment1.xml".into(),
                comments: vec![Comment::new(
                    1,
                    "This title slide introduces the typed PPTX feature graph.",
                    INCH,
                    INCH,
                )],
            },
            CommentList {
                slide_part_name: "/ppt/slides/slide2.xml".into(),
                relationship_id: "rIdComments2".into(),
                part_name: "/ppt/comments/comment2.xml".into(),
                comments: vec![
                    Comment::new(
                        1,
                        "The table is rendered as semantic writer row snapshots.",
                        INCH,
                        2 * INCH,
                    )
                    .with_index(2),
                ],
            },
            CommentList {
                slide_part_name: "/ppt/slides/slide6.xml".into(),
                relationship_id: "rIdComments6".into(),
                part_name: "/ppt/comments/comment3.xml".into(),
                comments: vec![
                    Comment::new(1, "First typed comment marker.", INCH, 7 * INCH / 2)
                        .with_index(3),
                    Comment::new(1, "Second typed comment marker.", 5 * INCH, 7 * INCH / 2)
                        .with_index(4),
                    Comment::new(1, "Third typed comment marker.", INCH, 5 * INCH).with_index(5),
                ],
            },
        ],
    };
    store_presentation_comments(opc, &comments, CommentConformance::Transitional)
}

fn store_media(opc: &mut OpcPackage, media: Vec<MediaAttachment>) -> Result<()> {
    let mut by_slide = BTreeMap::<usize, Vec<MediaAttachment>>::new();
    for attachment in media {
        by_slide
            .entry(attachment.slide)
            .or_default()
            .push(attachment);
    }

    let mut resource_index = 1usize;
    for (slide, attachments) in by_slide {
        let mut pictures = Vec::with_capacity(attachments.len());
        for (offset, attachment) in attachments.into_iter().enumerate() {
            let kind = match attachment.format.kind() {
                AuthoringMediaKind::Audio => MediaKind::Audio,
                AuthoringMediaKind::Video => MediaKind::Video,
            };
            let relationship_id = format!("rIdMedia{resource_index}");
            let resource = Resource::new(
                format!(
                    "/ppt/media/media{resource_index}.{}",
                    attachment.format.extension()
                ),
                attachment.format.mime_type(),
                attachment.data,
            );
            pictures.push(Picture {
                shape_id: 10_000 + slide as u32 * 100 + offset as u32,
                name: attachment.name,
                kind,
                relationship_id,
                resource: Some(resource),
                poster: None,
                transform: Some(Transform::emu(
                    attachment.x,
                    attachment.y,
                    attachment.width,
                    attachment.height,
                )?),
                office_extension: None,
            });
            resource_index += 1;
        }
        let slide_name =
            PackURI::new(format!("/ppt/slides/slide{slide}.xml")).map_err(Error::Uri)?;
        store_slide_media(
            opc,
            &slide_name,
            &MediaList { pictures },
            MediaConformance::Transitional,
        )?;
    }
    Ok(())
}

fn store_sections(opc: &mut OpcPackage, fragment: &str) -> Result<()> {
    let presentation_name = PackURI::new("/ppt/presentation.xml").map_err(Error::Uri)?;
    let current = std::str::from_utf8(opc.get_part(&presentation_name)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?;
    let closing = "</p:presentation>";
    let offset = current
        .rfind(closing)
        .ok_or_else(|| Error::Invalid("presentation root has no closing element".into()))?;
    let mut updated = String::with_capacity(current.len() + fragment.len());
    updated.push_str(&current[..offset]);
    updated.push_str(fragment);
    updated.push_str(&current[offset..]);
    opc.get_part_mut(&presentation_name)?
        .set_blob(updated.into_bytes());
    Ok(())
}
