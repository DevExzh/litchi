//! Paired timing of the record 0338 write paths: the unified call against
//! the hand-written leaf edit doing the same work, on the same snapshot, in
//! ABBA order. Outputs are compared outside timing.
//!
//! - root metadata `set / title=…` for DOCX, XLSX and PPTX;
//! - an XLSX cell `set` with `type`, a cell `add`, a plain cell `set` (the
//!   existing hot path, which the larger cell registry must not slow);
//! - an XLSX run `set`;
//! - an XLSX picture `set src`.
//!
//! Run: `cargo run --release -p litchi-ops --features docx,xlsx,pptx
//! --example p0_residue_bench`.
#[cfg(all(feature = "docx", feature = "xlsx", feature = "pptx"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::{NewNode, OpenOptions, Props, SourceMode, create, open_bytes_with};
    use litchi_xlsx::{Address, FontChange};

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
    let title = "Quarterly review";

    // `P0_PROFILE=docx-root|xlsx-run` loops one unified leg 2,000 times for
    // a profiler (callgrind).
    if let Ok(leg) = std::env::var("P0_PROFILE") {
        if leg == "docx-root" {
            let file = open_bytes_with(create(FileFormat::Docx)?.to_bytes()?.as_slice(), &eager)?;
            for _ in 0..2_000 {
                black_box(file.set("/", Props::new().with("title", title))?);
            }
        } else {
            let file = create(FileFormat::Xlsx)?.set(
                "/Sheet1/A1",
                Props::new()
                    .with("type", "richtext")
                    .with("runs", r#"[{"text":"Bold ","bold":true},{"text":"plain"}]"#),
            )?;
            let file = open_bytes_with(file.to_bytes()?.as_slice(), &eager)?;
            for _ in 0..2_000 {
                black_box(file.set("/Sheet1/A1/run[2]", Props::new().with("bold", true))?);
            }
        }
        return Ok(());
    }

    // Root metadata.
    {
        let bytes = create(FileFormat::Docx)?.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_docx::Package::from_bytes(&bytes)?;
        let unified = file.set("/", Props::new().with("title", title))?;
        assert_eq!(
            unified.metadata()?.unwrap_or_default().title.as_deref(),
            Some(title)
        );
        pair(
            "docx set / title",
            Box::new(|| {
                let mut edit = package.edit_metadata();
                edit.update(|m| m.title = Some(title.to_owned()));
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("title", title))?);
                Ok(())
            }),
        )?;
    }
    {
        let bytes = create(FileFormat::Xlsx)?.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_xlsx::Package::from_bytes(&bytes)?;
        pair(
            "xlsx set / title",
            Box::new(|| {
                let mut edit = package.edit_metadata()?;
                edit.update(|m| m.title = Some(title.to_owned()))?;
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("title", title))?);
                Ok(())
            }),
        )?;
        pair(
            "xlsx set /metadata title vs set / title",
            Box::new(|| {
                black_box(file.set("/metadata", Props::new().with("title", title))?);
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("title", title))?);
                Ok(())
            }),
        )?;
    }
    {
        let bytes = create(FileFormat::Pptx)?.to_bytes()?;
        let file = open_bytes_with(bytes.as_slice(), &eager)?;
        let package = litchi_pptx::Package::from_bytes(&bytes)?;
        pair(
            "pptx set / title",
            Box::new(|| {
                let mut edit = package.edit_metadata()?;
                edit.update(|m| m.title = Some(title.to_owned()))?;
                black_box(edit.commit()?.into_package());
                Ok(())
            }),
            Box::new(|| {
                black_box(file.set("/", Props::new().with("title", title))?);
                Ok(())
            }),
        )?;
    }

    // XLSX cells, runs and pictures on one workbook.
    let runs = r#"[{"text":"Bold ","bold":true},{"text":"plain"}]"#;
    let png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";
    let jpeg_bytes = [
        0xFF_u8, 0xD8, 0xFF, 0xE0, 0, 0x10, 0x4A, 0x46, 0x49, 0x46, 0,
    ];
    let jpeg = "data:image/jpeg;base64,/9j/4AAQSkZJRgA=";
    let seeded = create(FileFormat::Xlsx)?
        .set(
            "/Sheet1/A1",
            Props::new().with("type", "richtext").with("runs", runs),
        )?
        .add("/Sheet1", NewNode::new("picture").prop("src", png))?
        .file;
    let bytes = seeded.to_bytes()?;
    let file = open_bytes_with(bytes.as_slice(), &eager)?;
    let package = litchi_xlsx::Package::from_bytes(&bytes)?;
    let b2 = Address::at(1, 1)?;
    let c3 = Address::at(2, 2)?;
    let a1 = Address::at(0, 0)?;
    pair(
        "xlsx set cell value+type=number",
        Box::new(|| {
            let mut edit = package.edit()?;
            edit.sheet(0)?
                .expect("sheet")
                .set(b2, litchi_xlsx::Number::new("12.5")?)?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(file.set(
                "/Sheet1/B2",
                Props::new().with("value", "12.5").with("type", "number"),
            )?);
            Ok(())
        }),
    )?;
    pair(
        "xlsx set cell value (plain)",
        Box::new(|| {
            let mut edit = package.edit()?;
            edit.sheet(0)?.expect("sheet").set(b2, "Total")?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(file.set("/Sheet1/B2", Props::new().with("value", "Total"))?);
            Ok(())
        }),
    )?;
    pair(
        "xlsx add cell (ref, value)",
        Box::new(|| {
            let mut edit = package.edit()?;
            edit.sheet(0)?.expect("sheet").set(c3, "added")?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(
                file.add(
                    "/Sheet1",
                    NewNode::new("cell")
                        .prop("ref", "C3")
                        .prop("value", "added"),
                )?,
            );
            Ok(())
        }),
    )?;
    pair(
        "xlsx set run bold",
        Box::new(|| {
            let workbook = package.workbook();
            let mut rich = workbook
                .sheet(0)?
                .expect("sheet")
                .rich_text(a1)?
                .expect("runs");
            rich.run_mut(1)
                .expect("run")
                .apply_font(&FontChange::new().bold(Some(true)))?;
            let mut edit = package.edit()?;
            edit.sheet(0)?.expect("sheet").set_rich_text(a1, rich)?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(file.set("/Sheet1/A1/run[2]", Props::new().with("bold", true))?);
            Ok(())
        }),
    )?;
    pair(
        "xlsx set picture src (both decode the data URI)",
        Box::new(|| {
            use base64::Engine as _;
            let payload = jpeg.split_once(',').map_or("", |(_, payload)| payload);
            let bytes = base64::engine::general_purpose::STANDARD.decode(payload)?;
            assert_eq!(bytes, jpeg_bytes);
            let mut edit = package.edit_drawings()?;
            edit.replace_image(0, 0, bytes, "image/jpeg", "jpeg")?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(file.set("/Sheet1/picture[1]", Props::new().with("src", jpeg))?);
            Ok(())
        }),
    )?;
    Ok(())
}

#[cfg(not(all(feature = "docx", feature = "xlsx", feature = "pptx")))]
fn main() {}
