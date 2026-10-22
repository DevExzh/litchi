//! Paired OOXML metadata timing (record 0253).
//!
//! Set: the unified `File::set("/", {title})` against the leaf
//! `edit_metadata().update().commit()` on the same snapshot, for XLSX, PPTX
//! and XLSB, in ABBA order; outputs are compared outside timing.
//! Open: the source-backed open of a real workbook against the extra
//! application-properties read the unified open now performs.
//!
//! Run: `cargo run --release -p litchi-ops --all-features --example
//! ooxml_metadata_bench`.
#[cfg(all(feature = "xlsx", feature = "pptx", feature = "xlsb"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::{File, OpenOptions, Props, SourceMode, create, open_bytes, open_bytes_with};

    const ROUNDS: usize = 41;
    const WARMUP: usize = 5;

    fn median(mut values: Vec<u128>) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    fn report(name: &str, native: Vec<u128>, unified: Vec<u128>) {
        let native = median(native);
        let unified = median(unified);
        #[allow(clippy::cast_precision_loss, reason = "a printed ratio")]
        let ratio = unified as f64 / native as f64;
        println!("{name}: native {native} ns, unified {unified} ns, ratio {ratio:.3}");
    }

    for format in [FileFormat::Xlsx, FileFormat::Pptx, FileFormat::Xlsb] {
        let bytes = create(format)?.to_bytes()?;
        let eager = open_bytes_with(
            bytes.as_slice(),
            &OpenOptions::new().source_mode(SourceMode::Eager),
        )?;
        let source = open_bytes(bytes.as_slice())?;
        let title = "Quarterly review";
        let native_once: Box<dyn Fn() -> Result<(), Box<dyn std::error::Error>>> = match format {
            FileFormat::Xlsx => {
                let workbook = litchi_xlsx::Package::from_bytes(&bytes)?;
                Box::new(move || {
                    let mut edit = workbook.edit_metadata()?;
                    edit.update(|m| m.title = Some(title.to_owned()))?;
                    black_box(edit.commit()?.into_package());
                    Ok(())
                })
            },
            FileFormat::Pptx => {
                let package = litchi_pptx::Package::from_bytes(&bytes)?;
                Box::new(move || {
                    let mut edit = package.edit_metadata()?;
                    edit.update(|m| m.title = Some(title.to_owned()))?;
                    black_box(edit.commit()?.into_package());
                    Ok(())
                })
            },
            _ => {
                let package = litchi_xlsb::Package::from_bytes(&bytes)?;
                Box::new(move || {
                    let mut edit = package.edit_metadata()?;
                    edit.update(|m| m.title = Some(title.to_owned()))?;
                    black_box(edit.commit()?.into_package());
                    Ok(())
                })
            },
        };
        let unified_once = |file: &File| -> Result<File, litchi_ops::Error> {
            file.set("/metadata", Props::new().with("title", title))
        };
        // The unified result, checked outside timing.
        native_once()?;
        let produced = open_bytes(unified_once(&source)?.to_bytes()?)?;
        assert_eq!(
            produced.metadata()?.unwrap_or_default().title.as_deref(),
            Some(title)
        );

        for (owner, unified) in [("eager", &eager), ("source", &source)] {
            let mut native_ns = Vec::new();
            let mut unified_ns = Vec::new();
            for round in 0..ROUNDS + WARMUP {
                for route in [false, true, true, false] {
                    let start = Instant::now();
                    if route {
                        black_box(unified_once(unified)?);
                    } else {
                        native_once()?;
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
            report(
                &format!("{format:?} set /title ({owner} open)"),
                native_ns,
                unified_ns,
            );
        }
    }

    // First set on a fresh source-backed PPTX handle (record 0262): the
    // handle builds the eager package it commits metadata on once, so this
    // pairs open + first set against the leaf's open + edit.
    {
        let bytes = create(FileFormat::Pptx)?.to_bytes()?;
        let title = "Quarterly review";
        // `OOXML_METADATA_PROFILE=native|unified` loops one leg of this pair
        // (`OOXML_METADATA_LOOPS` times, default 1,000) for a profiler.
        if let Ok(leg) = std::env::var("OOXML_METADATA_PROFILE") {
            let loops: usize = match std::env::var("OOXML_METADATA_LOOPS") {
                Ok(loops) => loops.parse()?,
                Err(_) => 1_000,
            };
            for _ in 0..loops {
                if leg == "unified" {
                    let file = open_bytes(bytes.as_slice())?;
                    black_box(file.set("/metadata", Props::new().with("title", title))?);
                } else {
                    let package = litchi_pptx::Package::from_bytes(&bytes)?;
                    let mut edit = package.edit_metadata()?;
                    edit.update(|m| m.title = Some(title.to_owned()))?;
                    black_box(edit.commit()?.into_package());
                }
            }
            return Ok(());
        }
        let mut native_ns = Vec::new();
        let mut unified_ns = Vec::new();
        for round in 0..ROUNDS + WARMUP {
            for route in [false, true, true, false] {
                let start = Instant::now();
                if route {
                    let file = open_bytes(bytes.as_slice())?;
                    black_box(file.set("/metadata", Props::new().with("title", title))?);
                } else {
                    let package = litchi_pptx::Package::from_bytes(&bytes)?;
                    let mut edit = package.edit_metadata()?;
                    edit.update(|m| m.title = Some(title.to_owned()))?;
                    black_box(edit.commit()?.into_package());
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
        report(
            "Pptx open + first set /title (source open vs leaf eager open)",
            native_ns,
            unified_ns,
        );
    }

    // Open: what the application-properties read adds to a source-backed
    // workbook open.
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsx/ExcelTables.xlsx");
    let bytes = std::fs::read(path)?;
    let mut open_ns = Vec::new();
    let mut app_ns = Vec::new();
    for round in 0..ROUNDS + WARMUP {
        for route in [false, true, true, false] {
            let start = Instant::now();
            if route {
                let package = litchi_opc::SourceBackedPackage::from_vec(
                    bytes.clone(),
                    litchi_opc::SourceLimits::default(),
                    litchi_opc::SourceCacheLimits::default(),
                    None,
                )?;
                let started = Instant::now();
                black_box(litchi_ooxml_common::properties::app::read_source_backed(
                    &package,
                )?);
                if round >= WARMUP {
                    app_ns.push(started.elapsed().as_nanos());
                }
            } else {
                black_box(open_bytes(bytes.as_slice())?);
                if round >= WARMUP {
                    open_ns.push(start.elapsed().as_nanos());
                }
            }
        }
    }
    let open = median(open_ns);
    let app = median(app_ns);
    #[allow(clippy::cast_precision_loss, reason = "a printed share")]
    let share = app as f64 / open as f64;
    println!("xlsx open: {open} ns; app.xml read {app} ns ({share:.3} of open)");
    Ok(())
}

#[cfg(not(all(feature = "xlsx", feature = "pptx", feature = "xlsb")))]
fn main() {}
