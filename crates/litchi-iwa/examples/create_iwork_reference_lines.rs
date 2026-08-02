//! Create scratch Pages, Numbers, and Keynote charts with value-axis reference lines.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, ChartKind, ChartReferenceLine, ChartReferenceLineValue};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_reference_lines <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;
    let reference_lines = reference_lines()?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Reference Lines")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        ChartKind::Line2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 500.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Revenue thresholds")?;
    numbers.set_sheet_chart_reference_lines(
        sheet_id,
        chart.drawable_object_id,
        &reference_lines,
    )?;
    assert_eq!(
        numbers.sheet_chart_reference_lines(sheet_id, chart.drawable_object_id)?,
        reference_lines
    );
    numbers.save(output.join("reference-lines-crate.numbers"))?;

    let body = "Reference-line CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Line2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 500.0,
            height: 300.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Revenue thresholds")?;
    pages.set_body_chart_reference_lines(chart.drawable_object_id, &reference_lines)?;
    assert_eq!(
        pages.body_chart_reference_lines(chart.drawable_object_id)?,
        reference_lines
    );
    pages.save(output.join("reference-lines-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Reference-Line CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        ChartKind::Line2d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Revenue thresholds")?;
    keynote.set_slide_chart_reference_lines(0, chart.drawable_object_id, &reference_lines)?;
    assert_eq!(
        keynote.slide_chart_reference_lines(0, chart.drawable_object_id)?,
        reference_lines
    );
    keynote.save(output.join("reference-lines-crate.key"))?;

    println!("created reference-line fixtures in {}", output.display());
    Ok(())
}

fn reference_lines() -> Result<Vec<ChartReferenceLine>, Box<dyn std::error::Error>> {
    Ok(vec![
        ChartReferenceLine::minimum().with_name("Observed floor"),
        ChartReferenceLine::average().with_value_visibility(true),
        ChartReferenceLine::custom(ChartReferenceLineValue::new(30.0)?)
            .with_name("Target")
            .with_value_visibility(true),
    ])
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["Revenue".to_owned(), "Cost".to_owned()],
        (1..=12).map(|month| format!("M{month}")).collect(),
        vec![
            vec![
                Some(12.0),
                Some(18.0),
                Some(15.0),
                Some(21.0),
                Some(25.0),
                Some(22.0),
                Some(28.0),
                Some(31.0),
                Some(29.0),
                Some(35.0),
                Some(38.0),
                Some(42.0),
            ],
            vec![
                Some(8.0),
                Some(10.0),
                Some(9.0),
                Some(13.0),
                Some(14.0),
                Some(15.0),
                Some(17.0),
                Some(19.0),
                Some(18.0),
                Some(21.0),
                Some(23.0),
                Some(25.0),
            ],
        ],
    )?)
}
