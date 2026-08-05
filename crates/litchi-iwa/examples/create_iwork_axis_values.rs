//! Create compact Pages, Numbers, and Keynote fixtures for typed axis values.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{
    Axis, Bound, Bounds, ChartData, Kind, Direction, LabelAngle, MajorStepCount,
    MinorStepCount, Scale, Steps,
};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_axis_values <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let bounds = Bounds::fixed(Bound::new(1.0)?, Bound::new(30.0)?)?;
    let steps = Steps::fixed(MajorStepCount::new(6)?, MinorStepCount::new(2)?);

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Axis Values")
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
    numbers.set_sheet_chart_direction(sheet_id, chart.drawable_object_id, Direction::Columns)?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Axis values")?;
    set_numbers_axis_values(
        &mut numbers,
        sheet_id,
        chart.drawable_object_id,
        bounds,
        steps,
    )?;
    numbers.save(output.join("axis-values-crate.numbers"))?;

    let body = "Axis values";
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
    pages.set_body_chart_direction(chart.drawable_object_id, Direction::Columns)?;
    pages.set_body_chart_title(chart.drawable_object_id, "Axis values")?;
    set_pages_axis_values(&mut pages, chart.drawable_object_id, bounds, steps)?;
    pages.save(output.join("axis-values-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new().title("Axis Values").build()?;
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
    keynote.set_slide_chart_direction(0, chart.drawable_object_id, Direction::Columns)?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Axis values")?;
    set_keynote_axis_values(&mut keynote, chart.drawable_object_id, bounds, steps)?;
    keynote.save(output.join("axis-values-crate.key"))?;

    println!("created axis-value fixtures in {}", output.display());
    Ok(())
}

fn set_numbers_axis_values(
    editor: &mut litchi_iwa::numbers::NumbersEditor,
    sheet_id: u64,
    chart_id: u64,
    bounds: Bounds,
    steps: Steps,
) -> litchi_iwa::Result<()> {
    editor.set_sheet_chart_value_axis_bounds(sheet_id, chart_id, bounds)?;
    editor.set_sheet_chart_value_axis_scale(sheet_id, chart_id, Scale::Logarithmic)?;
    editor.set_sheet_chart_value_axis_steps(sheet_id, chart_id, steps)?;
    editor.set_sheet_chart_axis_label_angle(
        sheet_id,
        chart_id,
        Axis::Value,
        LabelAngle::RIGHT_DIAGONAL,
    )?;
    assert_eq!(
        editor.sheet_chart_value_axis_bounds(sheet_id, chart_id)?,
        bounds
    );
    assert_eq!(
        editor.sheet_chart_value_axis_scale(sheet_id, chart_id)?,
        Scale::Logarithmic
    );
    assert_eq!(
        editor.sheet_chart_value_axis_steps(sheet_id, chart_id)?,
        steps
    );
    Ok(())
}

fn set_pages_axis_values(
    editor: &mut litchi_iwa::pages::PagesEditor,
    chart_id: u64,
    bounds: Bounds,
    steps: Steps,
) -> litchi_iwa::Result<()> {
    editor.set_body_chart_value_axis_bounds(chart_id, bounds)?;
    editor.set_body_chart_value_axis_scale(chart_id, Scale::Logarithmic)?;
    editor.set_body_chart_value_axis_steps(chart_id, steps)?;
    editor.set_body_chart_axis_label_angle(chart_id, Axis::Value, LabelAngle::RIGHT_DIAGONAL)?;
    assert_eq!(editor.body_chart_value_axis_bounds(chart_id)?, bounds);
    assert_eq!(
        editor.body_chart_value_axis_scale(chart_id)?,
        Scale::Logarithmic
    );
    assert_eq!(editor.body_chart_value_axis_steps(chart_id)?, steps);
    Ok(())
}

fn set_keynote_axis_values(
    editor: &mut litchi_iwa::keynote::KeynoteEditor,
    chart_id: u64,
    bounds: Bounds,
    steps: Steps,
) -> litchi_iwa::Result<()> {
    editor.set_slide_chart_value_axis_bounds(0, chart_id, bounds)?;
    editor.set_slide_chart_value_axis_scale(0, chart_id, Scale::Logarithmic)?;
    editor.set_slide_chart_value_axis_steps(0, chart_id, steps)?;
    editor.set_slide_chart_axis_label_angle(
        0,
        chart_id,
        Axis::Value,
        LabelAngle::RIGHT_DIAGONAL,
    )?;
    assert_eq!(editor.slide_chart_value_axis_bounds(0, chart_id)?, bounds);
    assert_eq!(
        editor.slide_chart_value_axis_scale(0, chart_id)?,
        Scale::Logarithmic
    );
    assert_eq!(editor.slide_chart_value_axis_steps(0, chart_id)?, steps);
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["Revenue".to_owned(), "Cost".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?)
}
