//! Create Pages, Numbers, and Keynote files with typed pie leader-line visibility.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, ChartKind, ChartPieLeaderLineVisibility};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const LEADER_LINES: [ChartPieLeaderLineVisibility; 3] = [
    ChartPieLeaderLineVisibility::Hidden,
    ChartPieLeaderLineVisibility::Visible,
    ChartPieLeaderLineVisibility::Hidden,
];

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_pie_leader_lines <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Pie Leader-Line CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        ChartKind::Pie2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 360.0,
        },
    )?;
    numbers.set_sheet_chart_title(
        sheet_id,
        chart.drawable_object_id,
        "One visible leader line",
    )?;
    numbers.set_sheet_chart_pie_leader_line_visibilities(
        sheet_id,
        chart.drawable_object_id,
        &LEADER_LINES,
    )?;
    numbers.save(output.join("chart-pie-leader-lines-crate.numbers"))?;

    let body = "Pie Leader-Line CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Pie2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 360.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "One visible leader line")?;
    pages.set_body_chart_pie_leader_line_visibilities(chart.drawable_object_id, &LEADER_LINES)?;
    pages.save(output.join("chart-pie-leader-lines-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Pie Leader-Line CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        ChartKind::Pie2d,
        data()?,
        DrawablePoint { x: 470.0, y: 190.0 },
        DrawableSize {
            width: 900.0,
            height: 760.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "One visible leader line")?;
    keynote.set_slide_chart_pie_leader_line_visibilities(
        0,
        chart.drawable_object_id,
        &LEADER_LINES,
    )?;
    keynote.save(output.join("chart-pie-leader-lines-crate.key"))?;

    println!(
        "created typed pie leader-line fixtures in {}",
        output.display()
    );
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned(), "West".to_owned()],
        vec!["Revenue".to_owned()],
        vec![vec![Some(12.0)], vec![Some(18.0)], vec![Some(24.0)]],
    )?)
}
