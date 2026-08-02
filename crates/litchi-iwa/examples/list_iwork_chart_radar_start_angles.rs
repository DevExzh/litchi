//! List typed radar start angles in an iWork application package.

use std::env;
use std::path::Path;

use litchi_iwa::keynote::KeynoteEditor;
use litchi_iwa::numbers::NumbersEditor;
use litchi_iwa::pages::PagesEditor;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments.next().ok_or(
        "usage: list_iwork_chart_radar_start_angles <input.pages|input.numbers|input.key>",
    )?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    match Path::new(&input)
        .extension()
        .and_then(|value| value.to_str())
    {
        Some("numbers") => list_numbers(&input)?,
        Some("pages") => list_pages(&input)?,
        Some("key") => list_keynote(&input)?,
        extension => return Err(format!("unsupported iWork extension {extension:?}").into()),
    }
    Ok(())
}

fn list_numbers(input: &str) -> Result<(), Box<dyn std::error::Error>> {
    let editor = NumbersEditor::open(input)?;
    for sheet in editor.sheets()? {
        for chart in editor.sheet_charts(sheet.object_id)? {
            if chart.kind.supports_radar_start_angle() {
                println!(
                    "Numbers sheet={} chart={} kind={:?} start_angle={}°",
                    sheet.object_id,
                    chart.drawable_object_id,
                    chart.kind,
                    editor
                        .sheet_chart_radar_start_angle(sheet.object_id, chart.drawable_object_id)?
                        .degrees()
                );
            }
        }
    }
    Ok(())
}

fn list_pages(input: &str) -> Result<(), Box<dyn std::error::Error>> {
    let editor = PagesEditor::open(input)?;
    for chart in editor.body_charts()? {
        if chart.kind.supports_radar_start_angle() {
            println!(
                "Pages chart={} kind={:?} start_angle={}°",
                chart.drawable_object_id,
                chart.kind,
                editor
                    .body_chart_radar_start_angle(chart.drawable_object_id)?
                    .degrees()
            );
        }
    }
    Ok(())
}

fn list_keynote(input: &str) -> Result<(), Box<dyn std::error::Error>> {
    let editor = KeynoteEditor::open(input)?;
    for slide in editor.slides()? {
        for chart in editor.slide_charts(slide.index)? {
            if chart.kind.supports_radar_start_angle() {
                println!(
                    "Keynote slide={} chart={} kind={:?} start_angle={}°",
                    slide.index,
                    chart.drawable_object_id,
                    chart.kind,
                    editor
                        .slide_chart_radar_start_angle(slide.index, chart.drawable_object_id)?
                        .degrees()
                );
            }
        }
    }
    Ok(())
}
