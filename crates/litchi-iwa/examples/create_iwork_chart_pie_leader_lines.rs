//! Create Pages, Numbers, and Keynote files with typed pie leader-line visibility.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{
    ChartData, Kind, ChartPieLabelDistance, LabelVisibility, LeaderLineVisibility,
};
use litchi_iwa::keynote::{KeynoteDocumentBuilder, KeynoteEditor};
use litchi_iwa::numbers::{NumbersDocumentBuilder, NumbersEditor};
use litchi_iwa::pages::{PagesDocumentBuilder, PagesEditor};
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

const LEADER_LINES: [LeaderLineVisibility; 3] = [
    LeaderLineVisibility::Hidden,
    LeaderLineVisibility::Visible,
    LeaderLineVisibility::Hidden,
];
const LABEL_VISIBILITIES: [LabelVisibility; 3] = [LabelVisibility::ALL; 3];

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
    let label_distances = [
        ChartPieLabelDistance::from_percent(160.0)?,
        ChartPieLabelDistance::from_percent(160.0)?,
        ChartPieLabelDistance::from_percent(160.0)?,
    ];

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Pie Leader-Line CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].object_id;
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Pie2d,
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
    numbers.set_sheet_chart_pie_label_visibilities(
        sheet_id,
        chart.drawable_object_id,
        &LABEL_VISIBILITIES,
    )?;
    numbers.set_sheet_chart_pie_label_distances(
        sheet_id,
        chart.drawable_object_id,
        &label_distances,
    )?;
    let numbers_path = output.join("chart-pie-leader-lines-crate.numbers");
    numbers.save(&numbers_path)?;
    let reopened = NumbersEditor::from_bytes(&std::fs::read(&numbers_path)?)?;
    assert_eq!(
        reopened.sheet_chart_pie_leader_line_visibilities(sheet_id, chart.drawable_object_id)?,
        LEADER_LINES
    );
    assert_eq!(
        reopened.sheet_chart_pie_label_visibilities(sheet_id, chart.drawable_object_id)?,
        LABEL_VISIBILITIES
    );
    assert_eq!(
        reopened.sheet_chart_pie_label_distances(sheet_id, chart.drawable_object_id)?,
        label_distances
    );

    let body = "Pie Leader-Line CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Pie2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 360.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "One visible leader line")?;
    pages.set_body_chart_pie_leader_line_visibilities(chart.drawable_object_id, &LEADER_LINES)?;
    pages.set_body_chart_pie_label_visibilities(chart.drawable_object_id, &LABEL_VISIBILITIES)?;
    pages.set_body_chart_pie_label_distances(chart.drawable_object_id, &label_distances)?;
    let pages_path = output.join("chart-pie-leader-lines-crate.pages");
    pages.save(&pages_path)?;
    let reopened = PagesEditor::from_bytes(&std::fs::read(&pages_path)?)?;
    assert_eq!(
        reopened.body_chart_pie_leader_line_visibilities(chart.drawable_object_id)?,
        LEADER_LINES
    );
    assert_eq!(
        reopened.body_chart_pie_label_visibilities(chart.drawable_object_id)?,
        LABEL_VISIBILITIES
    );
    assert_eq!(
        reopened.body_chart_pie_label_distances(chart.drawable_object_id)?,
        label_distances
    );

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Pie Leader-Line CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Pie2d,
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
    keynote.set_slide_chart_pie_label_visibilities(
        0,
        chart.drawable_object_id,
        &LABEL_VISIBILITIES,
    )?;
    keynote.set_slide_chart_pie_label_distances(0, chart.drawable_object_id, &label_distances)?;
    let keynote_path = output.join("chart-pie-leader-lines-crate.key");
    keynote.save(&keynote_path)?;
    let reopened = KeynoteEditor::from_bytes(&std::fs::read(&keynote_path)?)?;
    assert_eq!(
        reopened.slide_chart_pie_leader_line_visibilities(0, chart.drawable_object_id)?,
        LEADER_LINES
    );
    assert_eq!(
        reopened.slide_chart_pie_label_visibilities(0, chart.drawable_object_id)?,
        LABEL_VISIBILITIES
    );
    assert_eq!(
        reopened.slide_chart_pie_label_distances(0, chart.drawable_object_id)?,
        label_distances
    );

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
