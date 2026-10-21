//! Paired XLSB metadata-set timing over a source-backed handle (record 0261).
//!
//! Legs, all setting one `title` on the same input:
//! - `leaf-eager`: `xlsb::Package::edit_metadata().update().commit()` on an
//!   eager package (the leaf leg of record 0253);
//! - `leaf-source`: the leaf source transaction,
//!   `SourceBackedWorkbook::edit_metadata().update().commit()`;
//! - `unified-source`: `File::set("/metadata", {title})` on a source-backed
//!   handle (the default open route);
//! - `unified-eager`: the same set on an eagerly opened handle.
//!
//! The `save:` pairs add publication to memory after the set
//! (`Package::to_bytes`, `SourceBackedWorkbook::write_to`, `File::to_bytes`),
//! where the source owner plans its source topology.
//!
//! Rounds run the legs in ABBA order (A = leaf, B = unified) per pair. Each
//! sample's result is dropped inside the timed region, as a caller would.
//! Outputs are checked outside timing.
//!
//! Run (release, pinned):
//! `taskset -c N cargo run --release -p litchi-ops --features xlsb --example
//! xlsb_metadata_source_bench -- [rounds]`.
#[cfg(feature = "xlsb")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::io::Cursor;
    use std::time::Instant;

    use litchi_ops::{File, OpenOptions, Props, SourceMode, open_bytes_with};
    use litchi_xlsb as xlsb;

    type Leg<'a> = Box<dyn Fn() -> Result<(), Box<dyn std::error::Error>> + 'a>;

    const WARMUP: usize = 5;
    let rounds: usize = std::env::args()
        .nth(1)
        .map(|value| value.parse())
        .transpose()?
        .unwrap_or(41);
    let title = "Quarterly review";

    fn median(mut values: Vec<u128>) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    fn paired(
        rounds: usize,
        leaf: &dyn Fn() -> Result<(), Box<dyn std::error::Error>>,
        unified: &dyn Fn() -> Result<(), Box<dyn std::error::Error>>,
    ) -> Result<(u128, u128), Box<dyn std::error::Error>> {
        let mut leaf_ns = Vec::new();
        let mut unified_ns = Vec::new();
        for round in 0..rounds + WARMUP {
            for route in [false, true, true, false] {
                let start = Instant::now();
                if route {
                    unified()?;
                } else {
                    leaf()?;
                }
                let elapsed = start.elapsed().as_nanos();
                if round >= WARMUP {
                    if route {
                        unified_ns.push(elapsed);
                    } else {
                        leaf_ns.push(elapsed);
                    }
                }
            }
        }
        Ok((median(leaf_ns), median(unified_ns)))
    }

    // Corpus: a fresh workbook, a real fixture, and a generated 8-sheet
    // workbook of 16,000 numeric cells.
    let mut corpus = Vec::new();
    corpus.push(("new", xlsb::Package::new()?.to_bytes()?));
    corpus.push((
        "Simple.xlsb",
        std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/ooxml/xlsb/Simple.xlsb"),
        )?,
    ));
    {
        let mut builder = xlsb::Builder::new();
        for sheet in 0..8 {
            let mut worksheet = xlsb::builder::Worksheet::new(format!("S{sheet}"));
            for row in 0..200_u32 {
                for column in 0..10_u32 {
                    worksheet.set_cell(row, column, f64::from(row * 10 + column));
                }
            }
            builder.add_worksheet(worksheet);
        }
        corpus.push(("8x2000 cells", builder.build()?.to_bytes()?));
    }

    for (name, bytes) in &corpus {
        let eager_package = xlsb::Package::from_bytes(bytes)?;
        let source_workbook = xlsb::SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone()))?;
        let source_file = open_bytes_with(bytes.as_slice(), &OpenOptions::new())?;
        let eager_file = open_bytes_with(
            bytes.as_slice(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )?;
        let props = Props::new().with("title", title);

        let leaf_eager: Leg<'_> = Box::new(|| {
            let mut edit = eager_package.edit_metadata()?;
            edit.update(|m| m.title = Some(title.to_owned()))?;
            black_box(edit.commit()?.into_package());
            Ok(())
        });
        let leaf_source: Leg<'_> = Box::new(|| {
            let mut edit = source_workbook.edit_metadata()?;
            edit.update(|m| m.title = Some(title.to_owned()))?;
            black_box(edit.commit()?.into_package());
            Ok(())
        });
        let unified = |file: &File| -> Leg<'_> {
            let file = file.clone();
            let props = props.clone();
            Box::new(move || {
                black_box(file.set("/metadata", props.clone())?);
                Ok(())
            })
        };
        let unified_source = unified(&source_file);
        let unified_eager = unified(&eager_file);
        // Set then publish to memory: the source owner plans its topology at
        // publication, the eager owner serializes its package.
        let leaf_eager_save: Leg<'_> = Box::new(|| {
            let mut edit = eager_package.edit_metadata()?;
            edit.update(|m| m.title = Some(title.to_owned()))?;
            black_box(edit.commit()?.into_package().to_bytes()?);
            Ok(())
        });
        let leaf_source_save: Leg<'_> = Box::new(|| {
            let mut edit = source_workbook.edit_metadata()?;
            edit.update(|m| m.title = Some(title.to_owned()))?;
            let mut bytes = Vec::new();
            edit.commit()?.into_package().write_to(&mut bytes)?;
            black_box(bytes);
            Ok(())
        });
        let unified_save = |file: &File| -> Leg<'_> {
            let file = file.clone();
            let props = props.clone();
            Box::new(move || {
                black_box(file.set("/metadata", props.clone())?.to_bytes()?);
                Ok(())
            })
        };
        let unified_source_save = unified_save(&source_file);
        let unified_eager_save = unified_save(&eager_file);

        // Outputs, outside timing: every leg stores the same title.
        let published = source_file.set("/metadata", props.clone())?.to_bytes()?;
        let reopened = open_bytes_with(published, &OpenOptions::new())?;
        assert_eq!(
            reopened.metadata()?.unwrap_or_default().title.as_deref(),
            Some(title)
        );

        for (label, leaf, unified) in [
            (
                "unified-source / leaf-source",
                &leaf_source,
                &unified_source,
            ),
            ("unified-source / leaf-eager", &leaf_eager, &unified_source),
            ("unified-eager / leaf-eager", &leaf_eager, &unified_eager),
            (
                "save: unified-source / leaf-source",
                &leaf_source_save,
                &unified_source_save,
            ),
            (
                "save: unified-source / leaf-eager",
                &leaf_eager_save,
                &unified_source_save,
            ),
            (
                "save: unified-eager / leaf-eager",
                &leaf_eager_save,
                &unified_eager_save,
            ),
        ] {
            let (leaf_ns, unified_ns) = paired(rounds, leaf.as_ref(), unified.as_ref())?;
            #[allow(clippy::cast_precision_loss, reason = "a printed ratio")]
            let ratio = unified_ns as f64 / leaf_ns as f64;
            println!(
                "{{\"corpus\":\"{name}\",\"pair\":\"{label}\",\"leaf_ns\":{leaf_ns},\"unified_ns\":{unified_ns},\"ratio\":{ratio:.4}}}"
            );
        }
    }
    Ok(())
}

#[cfg(not(feature = "xlsb"))]
fn main() {}
