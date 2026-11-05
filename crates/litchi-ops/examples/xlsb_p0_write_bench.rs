//! Paired XLSB P0 single-op writes (record 0261): the leaf call against the
//! unified `File::set`/`File::add` that wraps it.
//!
//! Operations, each one native commit or one unified call:
//! - `value`: a scalar number on `A1` (`cell_values::Edit::set_scalar`), the
//!   pre-existing write, as the reference for the owner boundary;
//! - `format`: `font.bold` + solid `fill` on `A1`
//!   (`cell_values::Edit::set_format`);
//! - `formula`: `SUM(A1:A2)` on a new cell (`Edit::set_formula_text`);
//! - `merge`: merge `C1:D2` (`structure::Edit::merge_cells`);
//! - `name`: add a workbook-scoped name (`structure::Edit::set_defined_name`).
//!
//! Owners: `eager` starts both legs from an eagerly opened owner; `source`
//! starts the unified leg from the default source-backed open and the leaf
//! leg materializes a positional source over the same bytes
//! (`SourceBackedWorkbook::from_read_at` + `to_workbook`, the route the
//! unified byte open uses) for every edit, as the unified edit boundary
//! does. Rounds run ABBA (A = leaf,
//! B = unified); each sample drops its result inside the timed region.
//! Results are checked outside timing.
//!
//! Run (release, pinned):
//! `taskset -c N cargo run --release -p litchi-ops --features xlsb --example
//! xlsb_p0_write_bench -- [rounds]`.
#[cfg(feature = "xlsb")]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::time::Instant;

    use litchi_ops::{File, NewNode, OpenOptions, Props, SourceMode, open_bytes_with};
    use litchi_xlsb as xlsb;
    use litchi_xlsb::cell_values::{Format, Reference};
    use litchi_xlsb::formula::Range;
    use litchi_xlsb::merged_cells::MergedCell;

    type Outcome = Result<(), Box<dyn std::error::Error>>;

    const WARMUP: usize = 5;
    let rounds: usize = std::env::args()
        .nth(1)
        .map(|value| value.parse())
        .transpose()?
        .unwrap_or(31);

    fn median(mut values: Vec<u128>) -> u128 {
        values.sort_unstable();
        values[values.len() / 2]
    }

    fn workbook(cells: u32) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
        let mut sheet = xlsb::builder::Worksheet::new("Data");
        for index in 0..cells {
            sheet.set_cell(index / 10, index % 10, f64::from(index) + 0.5);
        }
        let mut builder = xlsb::Builder::new();
        builder.add_worksheet(sheet);
        Ok(builder.build()?.to_bytes()?)
    }

    for cells in [10_u32, 1_000, 10_000] {
        let bytes = workbook(cells)?;
        for owner in ["eager", "source"] {
            let unified_file = open_bytes_with(
                bytes.as_slice(),
                &OpenOptions::new().source_mode(if owner == "eager" {
                    SourceMode::Eager
                } else {
                    SourceMode::Auto
                }),
            )?;
            let eager_package = xlsb::Package::from_bytes(&bytes)?;
            let source = xlsb::SourceBackedWorkbook::from_read_at(std::sync::Arc::new(
                litchi_core::OwnedSource::new(bytes.clone()),
            ))?;
            let leaf_package = || -> Result<xlsb::Package, Box<dyn std::error::Error>> {
                Ok(if owner == "eager" {
                    eager_package.clone()
                } else {
                    source.to_workbook()?.into_package()
                })
            };
            let mut format = Format::default();
            format.bold = Some(true);
            format.fill = Some(Some(0xFFFF_EE00));
            let a1 = Reference::new(0, 0)?;
            let target = Reference::new(cells / 10 + 2, 0)?;
            let target_path = format!("/Data/A{}", cells / 10 + 3);

            let operations: Vec<(
                &str,
                Box<dyn Fn() -> Outcome + '_>,
                Box<dyn Fn() -> Outcome + '_>,
            )> = vec![
                (
                    "value",
                    Box::new(|| {
                        let mut edit = leaf_package()?
                            .edit_cell_values(0_usize)?
                            .resolve()?
                            .expect("sheet");
                        edit.set_scalar(a1, litchi_xlsb::cell_values::Scalar::Number(7.0))?;
                        black_box(edit.commit()?.into_package());
                        Ok(())
                    }),
                    Box::new(|| {
                        black_box(unified_file.set("/Data/A1", Props::new().with("value", 7))?);
                        Ok(())
                    }),
                ),
                (
                    "format",
                    Box::new(|| {
                        let mut edit = leaf_package()?
                            .edit_cell_values(0_usize)?
                            .resolve()?
                            .expect("sheet");
                        edit.set_format(a1, &format)?;
                        black_box(edit.commit()?.into_package());
                        Ok(())
                    }),
                    Box::new(|| {
                        black_box(unified_file.set(
                            "/Data/A1",
                            Props::new().with("font.bold", true).with("fill", "FFEE00"),
                        )?);
                        Ok(())
                    }),
                ),
                (
                    "formula",
                    Box::new(|| {
                        let mut edit = leaf_package()?
                            .edit_cell_values(0_usize)?
                            .resolve()?
                            .expect("sheet");
                        edit.set_formula_text(target, "SUM(A1:A2)")?;
                        black_box(edit.commit()?.into_package());
                        Ok(())
                    }),
                    Box::new(|| {
                        black_box(unified_file.set(
                            target_path.as_str(),
                            Props::new().with("formula", "SUM(A1:A2)"),
                        )?);
                        Ok(())
                    }),
                ),
                (
                    "merge",
                    Box::new(|| {
                        let mut edit = leaf_package()?.edit()?;
                        edit.merge_cells(0, MergedCell::new(0, 1, 2, 3))?;
                        black_box(edit.commit()?.into_package());
                        Ok(())
                    }),
                    Box::new(|| {
                        black_box(
                            unified_file.set("/Data/C1:D2", Props::new().with("merge", true))?,
                        );
                        Ok(())
                    }),
                ),
                (
                    "name",
                    Box::new(|| {
                        let mut edit = leaf_package()?.edit()?;
                        edit.set_defined_name("Totals", None, 0, Range::new(0, 4, 1, 1)?)?;
                        black_box(edit.commit()?.into_package());
                        Ok(())
                    }),
                    Box::new(|| {
                        black_box(
                            unified_file.add(
                                "/",
                                NewNode::new("namedrange")
                                    .prop("name", "Totals")
                                    .prop("ref", "Data!$B$1:$B$5"),
                            )?,
                        );
                        Ok(())
                    }),
                ),
            ];

            // Outputs, outside timing.
            let formatted: File =
                unified_file.set("/Data/A1", Props::new().with("font.bold", true))?;
            assert!(formatted.to_bytes()? != bytes);

            for (name, leaf, unified) in &operations {
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
                let (leaf_ns, unified_ns) = (median(leaf_ns), median(unified_ns));
                #[allow(clippy::cast_precision_loss, reason = "a printed ratio")]
                let ratio = unified_ns as f64 / leaf_ns as f64;
                println!(
                    "{{\"cells\":{cells},\"owner\":\"{owner}\",\"op\":\"{name}\",\"leaf_ns\":{leaf_ns},\"unified_ns\":{unified_ns},\"ratio\":{ratio:.4}}}"
                );
            }
        }
    }
    Ok(())
}

#[cfg(not(feature = "xlsb"))]
fn main() {}
