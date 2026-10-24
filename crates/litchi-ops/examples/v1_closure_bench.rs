//! Paired timing of the record 0342 write paths: the unified call against
//! the hand-written leaf edit doing the same work, on the same snapshot, in
//! ABBA order. Outputs are compared outside timing.
//!
//! - DOCX `set / extended.company=…` (application properties) and
//!   `set / title=…` (the core-properties path, which now splits the core
//!   and application assignments);
//! - DOCX table `set border.top`, row `set header height.exact`;
//! - DOCX `add paragraph listStyle=bullet`;
//! - PPTX table-level `set` of the style switches.
//!
//! Run: `cargo run --release -p litchi-ops --features docx,pptx --example
//! v1_closure_bench`.
#[cfg(all(feature = "docx", feature = "pptx"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::{FileFormat, Position};
    use litchi_docx::document::{BodyAnchor, ElementPath, ElementStep, NewParagraph, NewRun};
    use litchi_docx::numbering::ListKind;
    use litchi_docx::properties::{
        Border, BorderStyle, EighthPoints, Height, HeightRule, HexColor, Twips,
    };
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

    let eager = OpenOptions::new().source_mode(SourceMode::Eager);

    // DOCX root metadata.
    {
        let bytes = create(FileFormat::Docx)?.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_docx::Package::from_bytes(&bytes)?;
        let unified = file.set("/", Props::new().with("extended.company", "Acme"))?;
        let mut edit = package.edit_metadata();
        let mut app = edit.app_props()?.unwrap_or_default();
        app.company = Some("Acme".to_owned());
        edit.set_app_props(Some(app));
        let native = edit.commit()?.into_package();
        assert_eq!(unified.to_bytes()?, native.to_bytes()?);
        pair(
            "docx set / extended.company",
            Box::new(|| {
                let mut edit = package.edit_metadata();
                let mut app = edit.app_props()?.unwrap_or_default();
                app.company = Some("Acme".to_owned());
                edit.set_app_props(Some(app));
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("extended.company", "Acme"))?);
                Ok(())
            }),
        )?;
        pair(
            "docx set / title",
            Box::new(|| {
                let mut edit = package.edit_metadata();
                edit.update(|metadata| metadata.title = Some("Quarterly review".to_owned()));
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("title", "Quarterly review"))?);
                Ok(())
            }),
        )?;
    }

    // DOCX table and row properties.
    {
        let table = create(FileFormat::Docx)?
            .add(
                "/body",
                NewNode::new("table").prop("rows", "3").prop("cols", "4"),
            )?
            .file;
        let bytes = table.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_docx::Package::from_bytes(&bytes)?;
        let path = ElementPath::new([ElementStep::Table(Position::new(0))]);
        let row = path.child(ElementStep::Row(Position::new(0)));
        let border = || {
            let mut border = Border::new(BorderStyle::from_xml("double").expect("style"));
            border.size = Some(EighthPoints(8));
            border.color = Some(HexColor::Rgb([0x1F, 0x38, 0x64]));
            border
        };
        let native_table = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.format_table_at(&path, |table| {
                table.borders.get_or_insert_default().top = Some(border());
            })?;
            Ok(edit.commit()?.into_package())
        };
        let unified = file.set(
            "/body/tbl[1]",
            Props::new().with("border.top", "double;8;1F3864"),
        )?;
        assert_eq!(unified.to_bytes()?, native_table()?.to_bytes()?);
        pair(
            "docx set table border.top",
            Box::new(|| {
                black_box(native_table()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set(
                    "/body/tbl[1]",
                    Props::new().with("border.top", "double;8;1F3864"),
                )?);
                Ok(())
            }),
        )?;
        let native_row = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.format_row_at(&row, |row| {
                row.header = Some(true);
                row.height = Some(Height::new(Twips(400), Some(HeightRule::Exact)));
            })?;
            Ok(edit.commit()?.into_package())
        };
        let row_props = || {
            Props::new()
                .with("header", true)
                .with("height.exact", "400")
        };
        let unified = file.set("/body/tbl[1]/tr[1]", row_props())?;
        assert_eq!(unified.to_bytes()?, native_row()?.to_bytes()?);
        pair(
            "docx resolve table vs row path (attribution)",
            Box::new(|| {
                black_box(file.get("/body/tbl[1]")?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.get("/body/tbl[1]/tr[1]")?);
                Ok(())
            }),
        )?;
        pair(
            "docx set row header height.exact",
            Box::new(|| {
                black_box(native_row()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/body/tbl[1]/tr[1]", row_props())?);
                Ok(())
            }),
        )?;
    }

    // DOCX paragraph add with a list style.
    {
        let bytes = create(FileFormat::Docx)?.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_docx::Package::from_bytes(&bytes)?;
        let native = || -> Result<litchi_docx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            let paragraph = NewParagraph::new().with_run(NewRun::new("Bulleted item"));
            let position = edit.insert_formatted_paragraph(BodyAnchor::End, &paragraph)?;
            let list = edit.add_list(ListKind::Bullet, None)?;
            edit.format_paragraph(position, |properties| {
                let numbering = properties.numbering.get_or_insert_default();
                numbering.id = Some(list.num_id());
                numbering.level = Some(0);
            })?;
            Ok(edit.commit()?.into_package())
        };
        let node = || {
            NewNode::new("paragraph")
                .prop("text", "Bulleted item")
                .prop("listStyle", "bullet")
        };
        let unified = file.add("/body", node())?.file;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "docx add paragraph listStyle=bullet",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.add("/body", node())?);
                Ok(())
            }),
        )?;
    }

    // PPTX table style switches.
    {
        let mut builder = litchi_pptx::builder::Builder::new()?;
        builder.presentation()?.add_slide()?;
        let package = builder.build()?;
        let mut edit = package.edit()?;
        edit.add_table(
            0usize,
            &[
                vec!["a1".to_owned(), "b1".to_owned()],
                vec!["a2".to_owned(), "b2".to_owned()],
            ],
            (0, 1_000_000, 2_000_000, 600_000),
        )?;
        let bytes = edit.commit()?.into_package().to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_pptx::Package::from_bytes(&bytes)?;
        let native = || -> Result<litchi_pptx::Package, Box<dyn std::error::Error>> {
            let mut edit = package.edit()?;
            edit.set_table_properties(0usize, 0usize, |properties| {
                properties.last_row = true;
                properties.first_col = true;
                properties.band_row = false;
            })?;
            Ok(edit.commit()?.into_package())
        };
        let props = || {
            Props::new()
                .with("lastRow", true)
                .with("firstCol", true)
                .with("bandedRows", false)
        };
        let unified = file.set("/slide[1]/table[1]", props())?;
        assert_eq!(unified.to_bytes()?, native()?.to_bytes()?);
        pair(
            "pptx set table switches",
            Box::new(|| {
                black_box(native()?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/slide[1]/table[1]", props())?);
                Ok(())
            }),
        )?;
    }
    Ok(())
}

#[cfg(not(all(feature = "docx", feature = "pptx")))]
fn main() {}
