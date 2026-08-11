//! Create scratch Pages, Numbers, and Keynote scatter charts with typed connections.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, ChartSeriesConnectionLine, Kind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_scatter_connections <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Scatter Connections")
        .build()?;
    let sheet_id = numbers.sheets()?[0].id();
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Scatter2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Scatter connections")?;
    numbers.set_sheet_chart_series_connection_lines(
        sheet_id,
        chart.drawable_object_id,
        &[ChartSeriesConnectionLine::Curved],
    )?;
    assert_eq!(
        numbers.sheet_chart_series_connection_lines(sheet_id, chart.drawable_object_id)?,
        vec![ChartSeriesConnectionLine::Curved]
    );
    numbers.save(output.join("scatter-curved-connection.numbers"))?;

    let body = "Scatter connection CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Scatter2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Scatter connections")?;
    pages.set_body_chart_series_connection_lines(
        chart.drawable_object_id,
        &[ChartSeriesConnectionLine::Curved],
    )?;
    assert_eq!(
        pages.body_chart_series_connection_lines(chart.drawable_object_id)?,
        vec![ChartSeriesConnectionLine::Curved]
    );
    pages.save(output.join("scatter-curved-connection.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Scatter Connection CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Scatter2d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Scatter connections")?;
    keynote.set_slide_chart_series_connection_lines(
        0,
        chart.drawable_object_id,
        &[ChartSeriesConnectionLine::Curved],
    )?;
    assert_eq!(
        keynote.slide_chart_series_connection_lines(0, chart.drawable_object_id)?,
        vec![ChartSeriesConnectionLine::Curved]
    );
    keynote.save(output.join("scatter-curved-connection.key"))?;

    println!("created scatter fixtures in {}", output.display());
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["X".to_owned(), "Y".to_owned()],
        vec!["1".to_owned(), "2".to_owned(), "4".to_owned()],
        vec![
            vec![Some(1.0), Some(2.0), Some(4.0)],
            vec![Some(4.0), Some(2.0), Some(3.0)],
        ],
    )?)
}
