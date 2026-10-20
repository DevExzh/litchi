//! Paired native-versus-unified probe of the record-0323 PPT elements.
//!
//! Every write scenario checks that both legs publish identical bytes before
//! timing; every read scenario that both legs read the same value. Samples
//! alternate in ABBA order on warm owners opened outside the timer: the
//! unified leg includes path parsing, resolution, registry admission, native
//! staging, publication and adoption; the native leg is the equivalent
//! hand-written `Package::edit` (or presentation read).
//!
//! ```sh
//! cargo build --release --offline --manifest-path tools/perf-baseline/Cargo.toml \
//!   --example ppt_elements_probe
//! taskset -c 7 tools/perf-baseline/target/release/examples/ppt_elements_probe 21 > elements.json
//! ```

use std::error::Error;
use std::time::Instant;

use litchi_ppt::Package;
use litchi_ppt::comments::CommentChange;
use litchi_ppt::hyperlink::{Click, Target as LinkTarget};
use litchi_ppt::package::edit::{Position, Target};
use litchi_ppt::text_edit::RunFormat;

type Result<T> = std::result::Result<T, Box<dyn Error>>;

/// A deck of `slides` slides; the first holds a text box, a 3x3 table and a
/// comment.
fn deck(slides: usize) -> Result<Vec<u8>> {
    use litchi_ppt::writer::{SlideComment, Table};
    let mut builder = litchi_ppt::writer::Builder::new();
    for index in 0..slides {
        let slide = builder.add_slide()?;
        builder.add_textbox(
            slide,
            24,
            24,
            400,
            48,
            &format!("Hello brave new world {index}"),
        )?;
        let mut table = Table::new(3, 3)?;
        for row in 0..3 {
            for column in 0..3 {
                table.set_cell_text(row, column, &format!("R{row}C{column}"))?;
            }
        }
        builder.add_table(slide, 24, 120, table)?;
        builder.add_comment(
            slide,
            SlideComment::new("Ada Lovelace", "Check the figures", 10, 20),
        )?;
    }
    Ok(builder.build()?.to_bytes()?)
}

fn median(values: &mut [f64]) -> f64 {
    values.sort_by(f64::total_cmp);
    let middle = values.len() / 2;
    if values.len().is_multiple_of(2) {
        (values[middle - 1] + values[middle]) / 2.0
    } else {
        values[middle]
    }
}

/// ABBA rounds of two legs; returns the scenario's JSON line.
fn paired(
    scenario: &str,
    slides: usize,
    rounds: usize,
    mut native: impl FnMut() -> Result<()>,
    mut unified: impl FnMut() -> Result<()>,
) -> Result<String> {
    native()?;
    unified()?;
    let mut native_us = Vec::with_capacity(rounds * 2);
    let mut unified_us = Vec::with_capacity(rounds * 2);
    let time = |leg: &mut dyn FnMut() -> Result<()>| -> Result<f64> {
        let start = Instant::now();
        leg()?;
        Ok(start.elapsed().as_secs_f64() * 1e6)
    };
    for _ in 0..rounds {
        native_us.push(time(&mut native)?);
        unified_us.push(time(&mut unified)?);
        unified_us.push(time(&mut unified)?);
        native_us.push(time(&mut native)?);
    }
    let mut ratios = native_us
        .iter()
        .zip(&unified_us)
        .map(|(native, unified)| unified / native)
        .collect::<Vec<_>>();
    let round = |values: &[f64]| {
        values
            .iter()
            .map(|value| (value * 1000.0).round() / 1000.0)
            .collect::<Vec<_>>()
    };
    let (raw_native, raw_unified) = (round(&native_us), round(&unified_us));
    Ok(format!(
        "{{\"scenario\":\"{scenario}\",\"slides\":{slides},\"samples\":{},\"native_median_us\":{:.3},\"unified_median_us\":{:.3},\"pair_ratio_median\":{:.4},\"native_us\":{raw_native:?},\"unified_us\":{raw_unified:?}}}",
        native_us.len(),
        median(&mut native_us),
        median(&mut unified_us),
        median(&mut ratios),
    ))
}

/// The writes: both legs publish identical bytes, checked once.
fn writes(rounds: usize, slides: usize, lines: &mut Vec<String>) -> Result<()> {
    let bytes = deck(slides)?;
    let native = Package::from_bytes(&bytes)?;
    let unified = litchi::open_bytes(bytes.clone())?;
    let shape = Target::new(Position::new(0), Position::new(0));
    let table = Target::new(Position::new(0), Position::new(1));

    macro_rules! scenario {
        ($name:literal, $native:expr, $unified:expr) => {{
            let native_leg = || -> Result<Package> { $native };
            let unified_leg = || -> Result<litchi::File> { $unified };
            if native_leg()?.to_bytes()? != unified_leg()?.to_bytes()? {
                return Err(concat!($name, ": legs publish different bytes").into());
            }
            lines.push(paired(
                $name,
                slides,
                rounds,
                || native_leg().map(drop),
                || unified_leg().map(drop),
            )?);
        }};
    }

    scenario!(
        "ppt_comment_set_v1",
        {
            let mut edit = native.edit()?;
            edit.set_comment(
                Position::new(0),
                Position::new(0),
                &CommentChange {
                    text: Some("Checked, with a longer note".into()),
                    ..CommentChange::default()
                },
            )?;
            Ok(edit.commit()?.into_package())
        },
        Ok(unified.set(
            "/slide[1]/comment[1]",
            litchi_ops::props! { "text" => "Checked, with a longer note" },
        )?)
    );
    scenario!(
        "ppt_comment_add_v1",
        {
            let mut edit = native.edit()?;
            edit.add_comment(
                Position::new(0),
                &litchi_ppt::comments::NewComment {
                    author: "Grace Hopper".into(),
                    initials: "GH".into(),
                    text: "New".into(),
                    date: None,
                    x: 0,
                    y: 0,
                },
            )?;
            Ok(edit.commit()?.into_package())
        },
        Ok(unified
            .add(
                "/slide[1]",
                litchi_ops::NewNode::new("comment")
                    .prop("author", "Grace Hopper")
                    .prop("text", "New"),
            )?
            .file)
    );
    scenario!(
        "ppt_cell_text_set_v1",
        {
            let mut edit = native.edit()?;
            edit.set_table_cell_text(table, Position::new(1), Position::new(2), "Changed")?;
            Ok(edit.commit()?.into_package())
        },
        Ok(unified.set(
            "/slide[1]/table[1]/tr[2]/tc[3]",
            litchi_ops::props! { "text" => "Changed" },
        )?)
    );
    scenario!(
        "ppt_hyperlink_add_v1",
        {
            let mut edit = native.edit()?;
            edit.set_shape_hyperlink(
                shape,
                Some(&Click {
                    target: LinkTarget::Url("https://example.test/probe".into()),
                    tooltip: None,
                }),
            )?;
            Ok(edit.commit()?.into_package())
        },
        Ok(unified
            .add(
                "/slide[1]/shape[1]",
                litchi_ops::NewNode::new("hyperlink").prop("link", "https://example.test/probe"),
            )?
            .file)
    );
    scenario!(
        "ppt_run_format_v1",
        {
            let mut edit = native.edit()?;
            let units = format!("Hello brave new world {}", 0)
                .encode_utf16()
                .count();
            edit.set_run_format(
                shape,
                0,
                units,
                &RunFormat {
                    bold: Some(true),
                    ..RunFormat::default()
                },
            )?;
            Ok(edit.commit()?.into_package())
        },
        Ok(unified.set(
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            litchi_ops::props! { "bold" => true },
        )?)
    );
    Ok(())
}

/// The reads: both legs read the same value, checked once.
fn reads(rounds: usize, slides: usize, lines: &mut Vec<String>) -> Result<()> {
    let bytes = deck(slides)?;
    let native = Package::from_bytes(&bytes)?;
    let presentation = native.presentation()?;
    let unified = litchi::open_bytes(bytes.clone())?;

    let native_comment = || -> Result<String> {
        let slide = presentation.slide_at(0)?.ok_or("slide")?;
        Ok(slide.comments()?.first().ok_or("comment")?.author.clone())
    };
    let unified_comment = || -> Result<String> {
        let node = unified.get("/slide[1]/comment[1]")?;
        let author = node
            .view()
            .format()?
            .find(|property| property.key() == "author")
            .map(|property| property.value().to_string())
            .ok_or("author")?;
        Ok(author)
    };
    if native_comment()? != unified_comment()? {
        return Err("comment reads differ".into());
    }
    lines.push(paired(
        "ppt_comment_get_v1",
        slides,
        rounds,
        || native_comment().map(drop),
        || unified_comment().map(drop),
    )?);

    let native_run = || -> Result<usize> {
        let slide = presentation.slide_at(0)?.ok_or("slide")?;
        let id = slide.shapes()?.first().ok_or("shape")?.id();
        Ok(slide
            .shape_text_runs(id)?
            .ok_or("runs")?
            .runs()
            .first()
            .ok_or("run")?
            .length)
    };
    let unified_run = || -> Result<usize> {
        let node = unified.get("/slide[1]/shape[1]/paragraph[1]/run[1]")?;
        Ok(node.view().text()?.ok_or("text")?.encode_utf16().count() + 1)
    };
    if native_run()? != unified_run()? {
        return Err("run reads differ".into());
    }
    lines.push(paired(
        "ppt_run_get_v1",
        slides,
        rounds,
        || native_run().map(drop),
        || unified_run().map(drop),
    )?);
    Ok(())
}

fn main() -> Result<()> {
    let rounds = std::env::args()
        .nth(1)
        .map_or(Ok(11), |value| value.parse())?;
    let mut lines = Vec::new();
    for slides in [1usize, 50] {
        writes(rounds, slides, &mut lines)?;
        reads(rounds, slides, &mut lines)?;
    }
    println!("[\n{}\n]", lines.join(",\n"));
    Ok(())
}
