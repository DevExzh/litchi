//! Create Pages, Numbers, and Keynote files with typed Radar series styles.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, ChartRadarSeriesStyle, Kind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_radar_series_styles <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Radar Series Style CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Radar2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 360.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Fill radar style")?;
    numbers.set_sheet_chart_radar_series_style(
        sheet_id,
        chart.drawable_object_id,
        ChartRadarSeriesStyle::Fill,
    )?;
    numbers.save(output.join("chart-radar-series-style-crate.numbers"))?;

    let body = "Radar Series Style CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Radar2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 360.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Stroke radar style")?;
    pages.set_body_chart_radar_series_style(
        chart.drawable_object_id,
        ChartRadarSeriesStyle::Stroke,
    )?;
    pages.save(output.join("chart-radar-series-style-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Radar Series Style CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Radar2d,
        data()?,
        DrawablePoint { x: 470.0, y: 190.0 },
        DrawableSize {
            width: 900.0,
            height: 760.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Fill and stroke radar style")?;
    keynote.set_slide_chart_radar_series_style(
        0,
        chart.drawable_object_id,
        ChartRadarSeriesStyle::FillAndStroke,
    )?;
    keynote.save(output.join("chart-radar-series-style-crate.key"))?;

    println!(
        "created typed Radar series-style fixtures in {}",
        output.display()
    );
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec![
            "Quality".to_owned(),
            "Speed".to_owned(),
            "Safety".to_owned(),
            "Value".to_owned(),
            "Support".to_owned(),
        ],
        vec![
            vec![Some(72.0), Some(88.0), Some(91.0), Some(67.0), Some(83.0)],
            vec![Some(85.0), Some(64.0), Some(78.0), Some(92.0), Some(70.0)],
        ],
    )?)
}
