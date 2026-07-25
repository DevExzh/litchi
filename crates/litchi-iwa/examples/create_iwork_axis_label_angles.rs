//! Create scratch Pages, Numbers, and Keynote charts with typed label angles.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartAxis, ChartAxisLabelAngle, ChartData, ChartKind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_axis_label_angles <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Axis Label Angles")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        ChartKind::Line2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Axis label angles")?;
    set_numbers_angles(&mut numbers, sheet_id, chart.drawable_object_id)?;
    numbers.save(output.join("axis-label-angles-crate.numbers"))?;

    let body = "Axis label angle CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Line2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Axis label angles")?;
    set_pages_angles(&mut pages, chart.drawable_object_id)?;
    pages.save(output.join("axis-label-angles-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Axis Label Angle CRUD")
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
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Axis label angles")?;
    set_keynote_angles(&mut keynote, chart.drawable_object_id)?;
    keynote.save(output.join("axis-label-angles-crate.key"))?;

    println!("created axis-label-angle fixtures in {}", output.display());
    Ok(())
}

fn set_numbers_angles(
    editor: &mut litchi_iwa::numbers::NumbersEditor,
    sheet_id: u64,
    chart_id: u64,
) -> litchi_iwa::Result<()> {
    for (axis, angle) in angles() {
        assert_eq!(
            editor.sheet_chart_axis_label_angle(sheet_id, chart_id, axis)?,
            ChartAxisLabelAngle::HORIZONTAL
        );
        editor.set_sheet_chart_axis_label_angle(sheet_id, chart_id, axis, angle)?;
        assert_eq!(
            editor.sheet_chart_axis_label_angle(sheet_id, chart_id, axis)?,
            angle
        );
    }
    Ok(())
}

fn set_pages_angles(
    editor: &mut litchi_iwa::pages::PagesEditor,
    chart_id: u64,
) -> litchi_iwa::Result<()> {
    for (axis, angle) in angles() {
        assert_eq!(
            editor.body_chart_axis_label_angle(chart_id, axis)?,
            ChartAxisLabelAngle::HORIZONTAL
        );
        editor.set_body_chart_axis_label_angle(chart_id, axis, angle)?;
        assert_eq!(editor.body_chart_axis_label_angle(chart_id, axis)?, angle);
    }
    Ok(())
}

fn set_keynote_angles(
    editor: &mut litchi_iwa::keynote::KeynoteEditor,
    chart_id: u64,
) -> litchi_iwa::Result<()> {
    for (axis, angle) in angles() {
        assert_eq!(
            editor.slide_chart_axis_label_angle(0, chart_id, axis)?,
            ChartAxisLabelAngle::HORIZONTAL
        );
        editor.set_slide_chart_axis_label_angle(0, chart_id, axis, angle)?;
        assert_eq!(
            editor.slide_chart_axis_label_angle(0, chart_id, axis)?,
            angle
        );
    }
    Ok(())
}

fn angles() -> [(ChartAxis, ChartAxisLabelAngle); 2] {
    [
        (ChartAxis::Category, ChartAxisLabelAngle::LEFT_DIAGONAL),
        (ChartAxis::Value, ChartAxisLabelAngle::RIGHT_DIAGONAL),
    ]
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
