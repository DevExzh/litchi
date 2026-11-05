//! `File::dump` against a read-only traversal baseline (record 0336).
//!
//! Each input is opened once. The baseline leg walks the whole node tree
//! and reads every node's `format` map and text, which is what `dump` reads;
//! the dump leg runs `File::dump`, which also filters each property through
//! the registry and builds the batch items. The legs alternate ABBA within
//! one process (baseline, dump, dump, baseline per round pair).
//!
//! Inputs (relative to `test-data/`): a 400-paragraph DOCX, a 75,770-cell
//! XLSX and a generated 200-slide PPTX (one title and one text shape per
//! slide). `DUMP_SCENARIO=<docx|xlsx|pptx>` runs one. Each prints one JSON
//! line: the median and minimum microseconds of both legs and the ratio of
//! the medians. Run pinned:
//! `taskset -c 7 cargo run --release -p litchi-ops --example dump_bench
//! --features docx,xlsx,pptx`.

#[cfg(all(feature = "docx", feature = "xlsx", feature = "pptx"))]
fn main() -> Result<(), Box<dyn std::error::Error>> {
    use std::hint::black_box;
    use std::path::Path;
    use std::time::Instant;

    use litchi_core::FileFormat;
    use litchi_ops::node::Node;
    use litchi_ops::{BatchMode, File};

    const ROUNDS: usize = 10;
    const WARMUP: usize = 1;

    fn walk(node: &Node<'_>, sink: &mut usize) -> litchi_ops::Result<()> {
        let view = node.view();
        for prop in view.format()? {
            *sink += prop.key().len() + prop.to_string().len();
        }
        *sink += view.text()?.map_or(0, |text| text.len());
        for child in node.children() {
            walk(&child?, sink)?;
        }
        Ok(())
    }

    fn traverse(file: &File) -> litchi_ops::Result<usize> {
        let mut sink = 0;
        walk(&file.root()?, &mut sink)?;
        Ok(sink)
    }

    fn median(values: &mut [f64]) -> f64 {
        values.sort_by(f64::total_cmp);
        values[values.len() / 2]
    }

    fn deck() -> litchi_ops::Result<File> {
        let mut items = Vec::new();
        for slide in 1..=200 {
            let mut item = litchi_ops::BatchItem::new("add");
            item.parent = Some("/".into());
            item.kind = Some("slide".into());
            item.props = Some(litchi_ops::Props::new().with("layout", "Title Only"));
            items.push(item);
            let mut title = litchi_ops::BatchItem::new("set");
            title.path = Some(format!("/slide[{slide}]/shape[1]").into());
            title.props = Some(litchi_ops::Props::new().with("text", format!("Slide {slide}")));
            items.push(title);
            let mut shape = litchi_ops::BatchItem::new("add");
            shape.parent = Some(format!("/slide[{slide}]").into());
            shape.kind = Some("shape".into());
            shape.props = Some(
                litchi_ops::Props::new()
                    .with("x", "2cm")
                    .with("y", "5cm")
                    .with("width", "10cm")
                    .with("height", "3cm")
                    .with("text", format!("Body text of slide {slide}")),
            );
            items.push(shape);
        }
        let mut file = litchi_ops::create(FileFormat::Pptx)?;
        for chunk in items.chunks(60) {
            file = file.batch(chunk, BatchMode::Atomic)?.into_file();
        }
        Ok(file)
    }

    let only = std::env::var("DUMP_SCENARIO").ok();
    let data = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let inputs: [(&str, Box<dyn Fn() -> litchi_ops::Result<File>>); 3] = [
        (
            "docx",
            Box::new(|| litchi_ops::open(data.join("ooxml/docx/drawing.docx"))),
        ),
        (
            "xlsx",
            Box::new(|| {
                litchi_ops::open(data.join("poi/test-data/spreadsheet/no_drawing_patriarch.xlsx"))
            }),
        ),
        ("pptx", Box::new(deck)),
    ];
    for (name, open) in &inputs {
        if only.as_deref().is_some_and(|only| only != *name) {
            continue;
        }
        let file = open()?;
        let items = file.dump()?.items().len();
        let mut base = Vec::new();
        let mut dump = Vec::new();
        for round in 0..WARMUP + ROUNDS {
            let order = if round % 2 == 0 {
                [false, true]
            } else {
                [true, false]
            };
            for is_dump in order {
                let start = Instant::now();
                if is_dump {
                    black_box(file.dump()?);
                } else {
                    black_box(traverse(&file)?);
                }
                let micros = start.elapsed().as_secs_f64() * 1e6;
                if round >= WARMUP {
                    if is_dump {
                        dump.push(micros);
                    } else {
                        base.push(micros);
                    }
                }
            }
        }
        let base_min = base.iter().copied().fold(f64::INFINITY, f64::min);
        let dump_min = dump.iter().copied().fold(f64::INFINITY, f64::min);
        let base_median = median(&mut base);
        let dump_median = median(&mut dump);
        println!(
            "{{\"scenario\":\"{name}\",\"items\":{items},\"traverse_median_us\":{base_median:.1},\"traverse_min_us\":{base_min:.1},\"dump_median_us\":{dump_median:.1},\"dump_min_us\":{dump_min:.1},\"ratio\":{:.3}}}",
            dump_median / base_median
        );
    }
    Ok(())
}

#[cfg(not(all(feature = "docx", feature = "xlsx", feature = "pptx")))]
fn main() {}
