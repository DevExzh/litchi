//! Paired timing of the record 0345 write paths: the unified call against
//! the hand-written leaf edit doing the same work, on the same snapshot, in
//! ABBA order, on a warm eager handle. Outputs are compared outside timing.
//!
//! - DOCX hyperlink `set tooltip`;
//! - DOCX `add cell` (with text) under a row;
//! - DOCX `add column` (width and text) under a table, and column
//!   `set width`;
//! - DOCX comment `add` replying to another (`parentId`), and comment
//!   `set done`.
//!
//! The fixture is `create()` with 200 paragraphs, a 4x4 table, a hyperlink
//! and a comment.
//!
//! Run: `cargo run --release -p litchi-ops --features docx --example
//! docx_p0_final_bench`.
#[cfg(feature = "docx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::{FileFormat, Position};
    use litchi_docx::document::{CommentChange, ElementPath, ElementStep, NewCell, NewComment};
    use litchi_ops::{NewNode, OpenOptions, Props, SourceMode, create, open_bytes_with};

    const ROUNDS: usize = 41;
    const WARMUP: usize = 5;

    fn median(mut values: Vec<u128>) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    type Leg<'a> = Box<dyn Fn() -> Result<(), Box<dyn std::error::Error>> + 'a>;

    fn pair(
        name: &str,
        native: Leg<'_>,
        unified: Leg<'_>,
    ) -> Result<(), Box<dyn std::error::Error>> {
        let mut native_ns = Vec::new();
        let mut unified_ns = Vec::new();
        for round in 0..ROUNDS + WARMUP {
            for route in [false, true, true, false] {
                let start = Instant::now();
                if route {
                    unified()?;
                } else {
                    native()?;
                }
                let elapsed = start.elapsed().as_nanos();
                if round >= WARMUP {
                    if route {
                        unified_ns.push(elapsed);
                    } else {
                        native_ns.push(elapsed);
                    }
                }
            }
        }
        let (native, unified) = (median(native_ns), median(unified_ns));
        #[allow(clippy::cast_precision_loss, reason = "a printed ratio")]
        let ratio = unified as f64 / native as f64;
        println!("{name}: native {native} ns, unified {unified} ns, ratio {ratio:.3}");
        Ok(())
    }

    // The fixture.
    let mut file = create(FileFormat::Docx)?;
    for index in 0..200 {
        file = file
            .add(
                "/body",
                NewNode::new("paragraph").prop("text", format!("Paragraph {index}")),
            )?
            .file;
    }
    file = file
        .add(
            "/body",
            NewNode::new("table").prop("rows", "4").prop("cols", "4"),
        )?
        .file;
    file = file
        .add(
            "/body/p[1]",
            NewNode::new("hyperlink")
                .prop("url", "https://example.test/")
                .prop("text", "link"),
        )?
        .file;
    file = file
        .add(
            "/body/p[2]",
            NewNode::new("comment")
                .prop("author", "Ada")
                .prop("text", "Root"),
        )?
        .file;
    let bytes = file.to_bytes()?;
    let eager = OpenOptions::new().source_mode(SourceMode::Eager);
    let file = open_bytes_with(bytes.as_slice(), &eager)?;
    let package = litchi_docx::Package::from_bytes(&bytes)?;
    let root_id = package.document()?.comments()?[0].id();
    // Warm the handle's body index.
    let _ = file.get("/body/tbl[1]/tr[1]")?;

    let table = ElementPath::new([ElementStep::Table(Position::new(0))]);
    let row = table.child(ElementStep::Row(Position::new(0)));
    let link = ElementPath::new([
        ElementStep::Paragraph(Position::new(0)),
        ElementStep::Hyperlink(Position::new(0)),
    ]);
    let second = ElementPath::new([ElementStep::Paragraph(Position::new(1))]);

    // Hyperlink tooltip.
    {
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.set_hyperlink_tooltip_at(&link, Some("Opens the report"))?;
            Ok(edit.commit()?.into_package())
        };
        let props = || Props::new().with("tooltip", "Opens the report");
        let unified = file.set("/body/p[1]/hyperlink[1]", props())?;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx set hyperlink tooltip",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/body/p[1]/hyperlink[1]", props())?);
                Ok(())
            }),
        )?;
    }

    // Cell insertion.
    {
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.insert_cell_at(&row, None, &NewCell::new().with_text("Total"))?;
            Ok(edit.commit()?.into_package())
        };
        let node = || NewNode::new("cell").prop("text", "Total");
        let unified = file.add("/body/tbl[1]/tr[1]", node())?.file;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx add cell",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.add("/body/tbl[1]/tr[1]", node())?);
                Ok(())
            }),
        )?;
    }

    // Column insertion and width.
    {
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.insert_column_at(&table, None, Some(2400), "Q1")?;
            Ok(edit.commit()?.into_package())
        };
        let node = || {
            NewNode::new("column")
                .prop("width", "2400")
                .prop("text", "Q1")
        };
        let unified = file.add("/body/tbl[1]", node())?.file;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx add column",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.add("/body/tbl[1]", node())?);
                Ok(())
            }),
        )?;
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.set_column_width_at(&table, Position::new(0), 1800)?;
            Ok(edit.commit()?.into_package())
        };
        let props = || Props::new().with("width", "1800");
        let unified = file.set("/body/tbl[1]/col[1]", props())?;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx set column width",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/body/tbl[1]/col[1]", props())?);
                Ok(())
            }),
        )?;
    }

    // Attribution: path resolution alone.
    if std::env::var_os("LITCHI_BENCH_ATTRIBUTION").is_some() {
        pair(
            "docx resolve /body/tbl[1] vs /body/tbl[1]/col[1]",
            Box::new(|| {
                black_box(file.get("/body/tbl[1]")?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.get("/body/tbl[1]/col[1]")?);
                Ok(())
            }),
        )?;
        pair(
            "docx leaf document() vs document().comments()",
            Box::new(|| {
                black_box(package.document()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(package.document()?.comments()?);
                Ok(())
            }),
        )?;
        // A comment's view now reads its commentsExtended entry and walks
        // the body for its anchor; the paragraph view is the reference.
        pair(
            "docx view format: /body/p[2] vs /comments/comment[1]",
            Box::new(|| {
                let node = file.get("/body/p[2]")?;
                black_box(node.view().format()?.count());
                Ok(())
            }),
            Box::new(|| {
                let node = file.get("/comments/comment[1]")?;
                black_box(node.view().format()?.count());
                Ok(())
            }),
        )?;
        let plain = || {
            NewNode::new("comment")
                .prop("author", "Grace")
                .prop("text", "Reply")
        };
        pair(
            "docx add plain comment",
            Box::new(|| {
                let mut edit = package.edit()?;
                edit.add_comment_at(&second, None, &NewComment::new("Grace", "Reply"))?;
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.add("/body/p[2]", plain())?);
                Ok(())
            }),
        )?;
    }

    // Comment reply and resolution.
    {
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.add_comment_at(
                &second,
                None,
                &NewComment::new("Grace", "Reply").with_parent(root_id),
            )?;
            Ok(edit.commit()?.into_package())
        };
        let node = || {
            NewNode::new("comment")
                .prop("author", "Grace")
                .prop("text", "Reply")
                .prop("parentId", i64::from(root_id))
        };
        let unified = file.add("/body/p[2]", node())?.file;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx add comment reply",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.add("/body/p[2]", node())?);
                Ok(())
            }),
        )?;
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.set_comment(root_id, &CommentChange::new().with_done(true))?;
            Ok(edit.commit()?.into_package())
        };
        let props = || Props::new().with("done", true);
        let unified = file.set("/comments/comment[1]", props())?;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx set comment done",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/comments/comment[1]", props())?);
                Ok(())
            }),
        )?;
    }
    Ok(())
}

#[cfg(not(feature = "docx"))]
fn main() {}
