//! Paired leaf-vs-unified measurements of the OOXML write follow-ups of
//! record 0325.
//!
//! The inputs are made here: a document with 50 body paragraphs, a header
//! of three paragraphs, ten bookmarks and ten `PAGE` fields; a deck of one
//! slide with 20 text boxes; a workbook with 40 notes, every other one
//! rich. Each scenario times the leaf call and the unified call doing the
//! same work on the same eager owner, alternating them ABBA within one
//! process (leaf, unified, unified, leaf per round pair), and checks
//! outside the clock that both legs publish the same bytes:
//!
//! - `story_paragraph_set`: `align=center` on `/header[1]/p[2]` (leaf:
//!   `edit_story` with `format_paragraph_at`);
//! - `story_text_set`: `text` of `/header[1]/p[2]` (leaf: `edit_story` with
//!   `replace_text_at`);
//! - `bookmark_rename`: `name` of `/bookmark[5]` (leaf: `fields` and
//!   `refresh_fields` with `Update::RenameBookmark`);
//! - `field_instruction_set`: `instruction` of `/field[5]` (leaf:
//!   `Update::Instruction`);
//! - `field_add`: a `PAGE` field at the end of `/body/p[30]` (leaf:
//!   `insert_field_at`);
//! - `levels_add`: an `abstractNum` with three `level<N>.*` levels (leaf:
//!   `add_abstract_num` of the same definition);
//! - `shape_link_set`: `link` of `/slide[1]/shape[10]` (leaf:
//!   `set_shape_hyperlink`);
//! - `note_runs_add`: a two-run note (leaf: `edit_comments().add_runs`);
//! - `note_font_set`: `font.bold` of `/Sheet1/comment[20]` (leaf:
//!   `set_runs` of one formatted run);
//! - `note_text_set`: `text` of `/Sheet1/comment[19]` (record 0320's
//!   plain set, for comparison; leaf: `set`);
//! - `rich_comment_get`: the 20th note, rich, and its format map (leaf:
//!   `Worksheet::comments`).
//!
//! `FOLLOWUP_SCENARIO=<name>` runs one scenario. Each prints one JSON line:
//! the median and minimum microseconds of both legs, the ratio of the
//! medians and whether both legs publish the same bytes. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example
//! ooxml_followups_bench --features docx,pptx,xlsx`.

#[cfg(all(feature = "docx", feature = "pptx", feature = "xlsx"))]
#[allow(
    clippy::too_many_lines,
    reason = "one measurement script, one scenario after another"
)]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::{FileFormat, Position};
    use litchi_docx::document::fields::Update;
    use litchi_docx::document::{ElementPath, ElementStep, StoryTarget};
    use litchi_ops::{NewNode, OpenOptions, SourceMode, props};

    const ROUNDS: usize = 20;
    const WARMUP: usize = 2;

    let only = std::env::var("FOLLOWUP_SCENARIO").ok();
    let wanted = |name: &str| only.as_deref().is_none_or(|only| only == name);
    let eager = |bytes: Vec<u8>| {
        litchi_ops::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
    };

    fn stats(mut samples: Vec<f64>) -> (f64, f64) {
        samples.sort_by(f64::total_cmp);
        (samples[samples.len() / 2], samples[0])
    }

    /// Microseconds per call of each leg, `repeat` calls per sample, ABBA.
    fn paired(
        repeat: u32,
        mut leaf: impl FnMut(),
        mut unified: impl FnMut(),
    ) -> (Vec<f64>, Vec<f64>) {
        let time = |run: &mut dyn FnMut()| {
            let started = Instant::now();
            for _ in 0..repeat {
                run();
            }
            started.elapsed().as_secs_f64() * 1e6 / f64::from(repeat)
        };
        let (mut a, mut b) = (Vec::new(), Vec::new());
        for round in 0..WARMUP + ROUNDS / 2 {
            let first = time(&mut leaf);
            let second = time(&mut unified);
            let third = time(&mut unified);
            let fourth = time(&mut leaf);
            if round >= WARMUP {
                a.extend([first, fourth]);
                b.extend([second, third]);
            }
        }
        (a, b)
    }

    fn report(name: &str, same: bool, (leaf, unified): (Vec<f64>, Vec<f64>)) {
        let (leaf_median, leaf_min) = stats(leaf);
        let (unified_median, unified_min) = stats(unified);
        println!(
            "{}",
            serde_json::json!({
                "scenario": name,
                "leaf_median_us": leaf_median,
                "leaf_min_us": leaf_min,
                "unified_median_us": unified_median,
                "unified_min_us": unified_min,
                "ratio": unified_median / leaf_median,
                "same_bytes": same,
            })
        );
    }

    // ------------------------------------------------------------------
    // DOCX
    // ------------------------------------------------------------------
    let mut docx = litchi_ops::create(FileFormat::Docx)?;
    for index in 0..50 {
        docx = docx
            .add(
                "/body",
                NewNode::new("paragraph").prop("text", format!("Paragraph {index}")),
            )?
            .file;
    }
    docx = docx
        .add("/", NewNode::new("header").prop("text", "Top"))?
        .file;
    for index in 0..2 {
        docx = docx
            .add(
                "/header[1]",
                NewNode::new("paragraph").prop("text", format!("Header {index}")),
            )?
            .file;
    }
    for index in 0..10 {
        let paragraph = format!("/body/p[{}]", index * 4 + 1);
        docx = docx
            .add(
                paragraph.as_str(),
                NewNode::new("bookmark").prop("name", format!("Mark{index}")),
            )?
            .file
            .add(paragraph.as_str(), NewNode::new("page"))?
            .file;
    }
    let docx_bytes = docx.to_bytes()?;
    let docx_file = eager(docx_bytes.clone())?;
    let docx_package = litchi_docx::Package::from_bytes(&docx_bytes)?;
    let header = docx_package.header_footer_entries()?[0].name().to_owned();
    let second = ElementPath::new([ElementStep::Paragraph(Position::new(1))]);

    if wanted("story_paragraph_set") {
        let leaf = || {
            let mut edit = docx_package.edit().expect("edit");
            edit.edit_story(StoryTarget::Part(&header), |story| {
                story
                    .format_paragraph_at(&second, |paragraph| {
                        paragraph.justification =
                            Some(litchi_docx::properties::Justification::Center);
                    })
                    .map(|_| ())
            })
            .expect("story");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .set("/header[1]/p[2]", props! { "align" => "center" })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "story_paragraph_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("story_text_set") {
        let leaf = || {
            let mut edit = docx_package.edit().expect("edit");
            edit.edit_story(StoryTarget::Part(&header), |story| {
                story.replace_text_at(&second, "Changed").map(|_| ())
            })
            .expect("story");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .set("/header[1]/p[2]", props! { "text" => "Changed" })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "story_text_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("bookmark_rename") {
        let leaf = || {
            let mut edit = docx_package.edit().expect("edit");
            let index = edit.fields().expect("fields");
            let bookmark = index
                .bookmarks()
                .iter()
                .position(|bookmark| bookmark.name() == "Mark4")
                .expect("bookmark");
            edit.refresh_fields(
                &index,
                &[Update::RenameBookmark {
                    bookmark,
                    name: "Renamed".to_owned(),
                }],
            )
            .expect("rename");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .set("/bookmark[5]", props! { "name" => "Renamed" })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "bookmark_rename",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("field_instruction_set") {
        let leaf = || {
            let mut edit = docx_package.edit().expect("edit");
            let index = edit.fields().expect("fields");
            edit.refresh_fields(
                &index,
                &[Update::Instruction {
                    field: 4,
                    instruction: " NUMPAGES ".to_owned(),
                }],
            )
            .expect("instruction");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .set("/field[5]", props! { "instruction" => "NUMPAGES" })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "field_instruction_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("field_add") {
        let paragraph = ElementPath::new([ElementStep::Paragraph(Position::new(29))]);
        let leaf = || {
            let mut edit = docx_package.edit().expect("edit");
            edit.insert_field_at(&paragraph, None, " PAGE ", Some("1"))
                .expect("field");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .add("/body/p[30]", NewNode::new("page"))
                .expect("add")
                .file
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "field_add",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("levels_add") {
        use litchi_docx::numbering::catalog::NewAbstractNum;
        let leaf = || {
            let mut definition = NewAbstractNum::new(false).with_id(500);
            for (ilvl, format, text) in [
                (0_u8, "upperRoman", "%1."),
                (1, "lowerLetter", "%2)"),
                (2, "bullet", "\u{25AA}"),
            ] {
                let level = definition
                    .level(ilvl)
                    .cloned()
                    .expect("level")
                    .with_format(format)
                    .with_text(text);
                definition = definition.with_level(level);
            }
            let mut edit = docx_package.edit().expect("edit");
            edit.add_abstract_num(&definition).expect("add");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            docx_file
                .add(
                    "/numbering",
                    NewNode::new("abstractNum")
                        .prop("id", 500)
                        .prop("level0.format", "upperRoman")
                        .prop("level0.text", "%1.")
                        .prop("level1.format", "lowerLetter")
                        .prop("level1.text", "%2)")
                        .prop("level2.format", "bullet"),
                )
                .expect("add")
                .file
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "levels_add",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }

    // ------------------------------------------------------------------
    // PPTX
    // ------------------------------------------------------------------
    let mut deck = litchi_ops::create(FileFormat::Pptx)?
        .add("/", NewNode::new("slide").prop("layout", "Blank"))?
        .file;
    for index in 0..20 {
        deck = deck
            .add(
                "/slide[1]",
                NewNode::new("shape")
                    .prop("text", format!("Box {index}"))
                    .prop("width", "2in")
                    .prop("height", "1in"),
            )?
            .file;
    }
    let deck_bytes = deck.to_bytes()?;
    let deck_file = eager(deck_bytes.clone())?;
    let deck_package = litchi_pptx::Package::from_bytes(&deck_bytes)?;
    if wanted("shape_link_set") {
        let link = litchi_pptx::hyperlinks::Link {
            target: litchi_pptx::hyperlinks::Target::Url("https://example.com/".to_owned()),
            tooltip: None,
        };
        let leaf = || {
            let mut edit = deck_package.edit().expect("edit");
            edit.set_shape_hyperlink(0_usize, 9_usize, Some(&link))
                .expect("link");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            deck_file
                .set(
                    "/slide[1]/shape[10]",
                    props! { "link" => "https://example.com/" },
                )
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "shape_link_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }

    // ------------------------------------------------------------------
    // XLSX
    // ------------------------------------------------------------------
    let mut book = litchi_ops::create(FileFormat::Xlsx)?;
    for index in 0..40 {
        let node = NewNode::new("comment").prop("ref", format!("B{}", index + 2));
        let node = if index % 2 == 1 {
            node.prop(
                "runs",
                format!(r#"[{{"text":"Ann:","bold":true}},{{"text":" note {index}"}}]"#),
            )
        } else {
            node.prop("text", format!("note {index}"))
        };
        book = book.add("/Sheet1", node)?.file;
    }
    let book_bytes = book.to_bytes()?;
    let book_file = eager(book_bytes.clone())?;
    let book_package = litchi_xlsx::Package::from_bytes(&book_bytes)?;
    let note_run = |text: &str, bold: bool| {
        use litchi_xlsx::workbook::comments::{RunColor, RunFont, TextRun};
        TextRun {
            text: text.to_owned(),
            font: RunFont {
                bold: bold.then_some(true),
                size: Some(9.0),
                color: Some(RunColor::Indexed(81)),
                name: Some("Tahoma".to_owned()),
                ..RunFont::default()
            },
        }
    };
    if wanted("note_runs_add") {
        let runs = [note_run("Ann:", true), note_run(" check", false)];
        let leaf = || {
            let mut edit = book_package.edit_comments().expect("edit");
            edit.add_runs(0, "Z99", "Author", &runs).expect("add");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            book_file
                .add(
                    "/Sheet1",
                    NewNode::new("comment")
                        .prop("ref", "Z99")
                        .prop("runs", r#"[{"text":"Ann:","bold":true},{"text":" check"}]"#),
                )
                .expect("add")
                .file
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "note_runs_add",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("note_font_set") {
        let runs = [note_run("note 18", true)];
        let leaf = || {
            let mut edit = book_package.edit_comments().expect("edit");
            edit.set_runs(0, "B20", "Author", &runs).expect("set");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            book_file
                .set("/Sheet1/comment[19]", props! { "font.bold" => true })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "note_font_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    // The plain-text note set of record 0320, for comparison: its unified
    // leg resolves `comment[N]` (one parse of the comments part) as the
    // `font.*` set does.
    if wanted("note_text_set") {
        let leaf = || {
            let mut edit = book_package.edit_comments().expect("edit");
            edit.set(0, "B20", "Author", "changed").expect("set");
            edit.commit().expect("commit").into_package()
        };
        let unified = || {
            book_file
                .set("/Sheet1/comment[19]", props! { "text" => "changed" })
                .expect("set")
        };
        let same = leaf().to_bytes()? == unified().to_bytes()?;
        report(
            "note_text_set",
            same,
            paired(5, || drop(black_box(leaf())), || drop(black_box(unified()))),
        );
    }
    if wanted("rich_comment_get") {
        report(
            "rich_comment_get",
            true,
            paired(
                20,
                || {
                    let sheet = book_package
                        .workbook()
                        .sheet(0)
                        .expect("sheet")
                        .expect("Sheet1");
                    let comments = sheet.comments().expect("comments").expect("part");
                    black_box(comments.get("B21"));
                },
                || {
                    let node = book_file.get("/Sheet1/comment[20]").expect("comment");
                    black_box(node.view().format().expect("format").count());
                },
            ),
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "docx", feature = "pptx", feature = "xlsx")))]
fn main() {
    eprintln!("enable the `docx`, `pptx` and `xlsx` features");
}
