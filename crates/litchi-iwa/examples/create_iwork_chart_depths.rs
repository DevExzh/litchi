//! Create Pages, Numbers, and Keynote files with typed 3D chart depths.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{Chart3dDepth, ChartData, Kind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_depths <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let numbers_depth = Chart3dDepth::from_percent(25.0)?;
    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("3D Chart Depth CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Bar3d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "25% depth")?;
    numbers.set_sheet_chart_3d_depth(sheet_id, chart.drawable_object_id, numbers_depth)?;
    assert_eq!(
        numbers.sheet_chart_3d_depth(sheet_id, chart.drawable_object_id)?,
        numbers_depth
    );
    numbers.save(output.join("chart-depth-crate.numbers"))?;

    let pages_depth = Chart3dDepth::from_percent(50.0)?;
    let body = "3D Chart Depth CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Column3d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "50% depth")?;
    pages.set_body_chart_3d_depth(chart.drawable_object_id, pages_depth)?;
    assert_eq!(
        pages.body_chart_3d_depth(chart.drawable_object_id)?,
        pages_depth
    );
    pages.save(output.join("chart-depth-crate.pages"))?;

    let keynote_depth = Chart3dDepth::from_percent(75.0)?;
    let mut keynote = KeynoteDocumentBuilder::new()
        .title("3D Chart Depth CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Area3d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "75% depth")?;
    keynote.set_slide_chart_3d_depth(0, chart.drawable_object_id, keynote_depth)?;
    assert_eq!(
        keynote.slide_chart_3d_depth(0, chart.drawable_object_id)?,
        keynote_depth
    );
    keynote.save(output.join("chart-depth-crate.key"))?;

    println!("created typed chart-depth fixtures in {}", output.display());
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?)
}
