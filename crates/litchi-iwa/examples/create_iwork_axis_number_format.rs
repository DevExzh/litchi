//! Create scratch Pages, Numbers, and Keynote charts with typed axis formats.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{Axis, ChartData, Kind, DecimalPlaces, NegativeStyle, NumberFormat};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_axis_number_format <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;
    let format = axis_format()?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Axis Number Format")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Axis number format")?;
    assert_eq!(
        numbers.sheet_chart_axis_number_format(sheet_id, chart.drawable_object_id, Axis::Value,)?,
        NumberFormat::AXIS_NATIVE_DEFAULT
    );
    numbers.set_sheet_chart_axis_number_format(
        sheet_id,
        chart.drawable_object_id,
        Axis::Value,
        format,
    )?;
    assert_eq!(
        numbers.sheet_chart_axis_number_format(sheet_id, chart.drawable_object_id, Axis::Value,)?,
        format
    );
    numbers.save(output.join("axis-number-format-crate.numbers"))?;

    let body = "Axis number format CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Axis number format")?;
    assert_eq!(
        pages.body_chart_axis_number_format(chart.drawable_object_id, Axis::Value)?,
        NumberFormat::AXIS_NATIVE_DEFAULT
    );
    pages.set_body_chart_axis_number_format(chart.drawable_object_id, Axis::Value, format)?;
    assert_eq!(
        pages.body_chart_axis_number_format(chart.drawable_object_id, Axis::Value)?,
        format
    );
    pages.save(output.join("axis-number-format-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Axis Number Format CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Axis number format")?;
    assert_eq!(
        keynote.slide_chart_axis_number_format(0, chart.drawable_object_id, Axis::Value)?,
        NumberFormat::AXIS_NATIVE_DEFAULT
    );
    keynote.set_slide_chart_axis_number_format(0, chart.drawable_object_id, Axis::Value, format)?;
    assert_eq!(
        keynote.slide_chart_axis_number_format(0, chart.drawable_object_id, Axis::Value)?,
        format
    );
    keynote.save(output.join("axis-number-format-crate.key"))?;

    println!("created axis-format fixtures in {}", output.display());
    Ok(())
}

fn axis_format() -> Result<NumberFormat, Box<dyn std::error::Error>> {
    Ok(NumberFormat::new(
        DecimalPlaces::fixed(2)?,
        NegativeStyle::Parentheses,
        true,
    ))
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["Revenue".to_owned(), "Cost".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12_345.5), Some(-6_789.25), Some(9_876.75)],
            vec![Some(4_321.0), Some(5_432.5), Some(-3_210.75)],
        ],
    )?)
}
