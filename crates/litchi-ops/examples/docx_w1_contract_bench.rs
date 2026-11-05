//! Paired timing of the DOCX routes record 0351 §5 lists for wave W1
//! (record 0355): open, one paragraph set, one paragraph add, a metadata
//! title set, a 1,000-paragraph batch, save, and a repeated save of one
//! package. Each pair times the unified call
//! against the hand-written leaf call doing the same work, in ABBA order on
//! the same bytes; outputs are compared outside timing.
//!
//! The leaf legs use only API that exists before and after the W1 change
//! (`Package::from_bytes`, `edit`, `commit`, `to_bytes` through a
//! `&mut Package`), so the same source builds at the wave's base and head
//! for a before/after comparison of the leaf routes themselves.
//!
//! The fixture is `create()` with 1,000 paragraphs, opened with the default
//! (source-backed) owner; one batch leg also runs on the eager owner.
//!
//! Run: `cargo run --release -p litchi-ops --features docx --example
//! docx_w1_contract_bench`.
#[cfg(feature = "docx")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::{FileFormat, Position};
    use litchi_ops::{BatchItem, BatchMode, NewNode, Props, create, open_bytes, props};

    const ROUNDS: usize = 31;
    const WARMUP: usize = 5;
    const PARAGRAPHS: usize = 1_000;

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
        println!("docx {name}: native {native} ns, unified {unified} ns, ratio {ratio:.3}");
        Ok(())
    }

    /// Serialize through a `&mut Package` so the leg compiles against both the
    /// pre-W1 (`&mut self`) and the W1 (`&self`) output signature.
    fn bytes(package: &mut litchi_docx::Package) -> Result<Vec<u8>, litchi_docx::Error> {
        package.to_bytes()
    }

    // The fixture.
    let mut file = create(FileFormat::Docx)?;
    for index in 0..PARAGRAPHS {
        file = file
            .add(
                "/body",
                NewNode::new("paragraph").prop("text", format!("Paragraph {index}")),
            )?
            .file;
    }
    let source = file.to_bytes()?;
    let package = litchi_docx::Package::from_bytes(&source)?;
    let unified = open_bytes(source.clone())?;

    // Outputs agree before timing.
    let mut native_set = {
        let mut edit = package.edit()?;
        edit.replace_paragraph_text(Position::new(0), "changed")?;
        edit.commit()?.into_package()
    };
    let unified_set = unified.set("/body/p[1]", props! { "text" => "changed" })?;
    assert_eq!(
        litchi_docx::Package::from_bytes(&unified_set.to_bytes()?)?
            .document()?
            .text()?,
        litchi_docx::Package::from_bytes(&bytes(&mut native_set)?)?
            .document()?
            .text()?
    );

    pair(
        "open (1,000 paragraphs)",
        Box::new(|| {
            black_box(litchi_docx::Package::from_bytes(black_box(&source))?);
            Ok(())
        }),
        Box::new(|| {
            black_box(open_bytes(black_box(source.clone()))?);
            Ok(())
        }),
    )?;

    pair(
        "set paragraph text",
        Box::new(|| {
            let mut edit = package.edit()?;
            edit.replace_paragraph_text(Position::new(500), "changed")?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(unified.set("/body/p[501]", props! { "text" => "changed" })?);
            Ok(())
        }),
    )?;

    pair(
        "add paragraph",
        Box::new(|| {
            let mut edit = package.edit()?;
            edit.insert_paragraph(Position::new(PARAGRAPHS), "added")?;
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(unified.add("/body", NewNode::new("paragraph").prop("text", "added"))?);
            Ok(())
        }),
    )?;

    pair(
        "set / title",
        Box::new(|| {
            let mut edit = package.edit_metadata();
            edit.update(|props| props.title = Some("Title".to_owned()));
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(unified.set("/", props! { "title" => "Title" })?);
            Ok(())
        }),
    )?;

    pair(
        "File::edit 1k paragraph text sets",
        Box::new(|| {
            let mut edit = package.edit()?;
            for index in 0..PARAGRAPHS {
                edit.replace_paragraph_text(Position::new(index), "batched")?;
            }
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(unified.edit(|tx| {
                for index in 1..=PARAGRAPHS {
                    tx.set(
                        format!("/body/p[{index}]").as_str(),
                        props! { "text" => "batched" },
                    )?;
                }
                Ok(())
            })?);
            Ok(())
        }),
    )?;

    let items: Vec<BatchItem<'static>> = (1..=PARAGRAPHS)
        .map(|index| {
            let mut item = BatchItem::new("set");
            item.path = Some(format!("/body/p[{index}]").into());
            item.props = Some(Props::new().with("text", "batched"));
            item
        })
        .collect();
    pair(
        "File::batch 1k paragraph text sets",
        Box::new(|| {
            let mut edit = package.edit()?;
            for index in 0..PARAGRAPHS {
                edit.replace_paragraph_text(Position::new(index), "batched")?;
            }
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(unified.batch(items.clone(), BatchMode::Atomic)?);
            Ok(())
        }),
    )?;

    let eager = litchi_ops::open_bytes_with(
        source.clone(),
        &litchi_ops::OpenOptions::new().source_mode(litchi_ops::SourceMode::Eager),
    )?;
    pair(
        "File::batch 1k paragraph text sets (eager owner)",
        Box::new(|| {
            let mut edit = package.edit()?;
            for index in 0..PARAGRAPHS {
                edit.replace_paragraph_text(Position::new(index), "batched")?;
            }
            black_box(edit.commit()?.into_package());
            Ok(())
        }),
        Box::new(|| {
            black_box(eager.batch(items.clone(), BatchMode::Atomic)?);
            Ok(())
        }),
    )?;

    pair(
        "save (to_bytes) after one set",
        Box::new(|| {
            let mut copy = native_set.clone();
            black_box(bytes(&mut copy)?);
            Ok(())
        }),
        Box::new(|| {
            black_box(unified_set.to_bytes()?);
            Ok(())
        }),
    )?;

    pair(
        "repeated save of one package (second to_bytes)",
        Box::new(|| {
            let mut copy = native_set.clone();
            let first = bytes(&mut copy)?;
            let second = bytes(&mut copy)?;
            assert_eq!(first, second, "repeated output must be byte-identical");
            black_box(second);
            Ok(())
        }),
        Box::new(|| {
            let first = unified_set.to_bytes()?;
            let second = unified_set.to_bytes()?;
            assert_eq!(first, second, "repeated output must be byte-identical");
            black_box(second);
            Ok(())
        }),
    )?;
    Ok(())
}

#[cfg(not(feature = "docx"))]
fn main() {}
