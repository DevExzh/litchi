use std::error::Error;
use std::path::Path;

use xlsx0683_selected_record_probe::{benchmark, differential};

fn usage() -> &'static str {
    "usage:\n  xlsx0683 layout\n  xlsx0683 diff FILE.xlsx SHEET\n  xlsx0683 bench OP FILE.xlsx SHEET WARMUP SAMPLES\n\nOP is visit-cold, cells-cold, visit-selected, cells-selected, visit-warm, or cells-warm."
}

fn main() -> Result<(), Box<dyn Error>> {
    let arguments: Vec<String> = std::env::args().skip(1).collect();
    match arguments.first().map(String::as_str) {
        Some("scan") if arguments.len() == 2 => {
            use litchi_xlsx::raw::selected_worksheet::{RangeScanOutcome, scan_range};
            let bytes = std::fs::read(&arguments[1])?;
            let result = scan_range(
                &mut std::io::Cursor::new(bytes),
                &litchi_ooxml_common::mce::Capabilities::default(),
                &litchi_ooxml_common::mce::StreamLimits::default(),
                litchi_xlsx::Rect::ALL,
            );
            match result {
                Ok(RangeScanOutcome::Eligible(cells)) => println!("eligible:{}", cells.cells.len()),
                Ok(RangeScanOutcome::NotEligible(reason)) => println!("not-eligible:{reason:?}"),
                Err(error) => println!("refused:{error:?}"),
            }
        },
        Some("layout") => println!(
            "selected_record_bytes={}\tcell_bytes={}",
            std::mem::size_of::<litchi_xlsx::raw::selected_worksheet::SelectedRecord>(),
            std::mem::size_of::<litchi_xlsx::Cell>()
        ),
        Some("diff") if arguments.len() == 3 => {
            differential(Path::new(&arguments[1]), &arguments[2])?;
        },
        Some("bench") if arguments.len() == 6 => {
            let samples = benchmark(
                &arguments[1],
                Path::new(&arguments[2]),
                &arguments[3],
                arguments[4].parse()?,
                arguments[5].parse()?,
            )?;
            for (index, sample) in samples.iter().enumerate() {
                println!(
                    "sample={}\tnanos={}\tcallbacks={}\tresult={}",
                    index,
                    sample.nanos,
                    sample.callbacks,
                    if sample.succeeded { "ok" } else { "refused" }
                );
            }
        },
        _ => return Err(usage().into()),
    }
    Ok(())
}
