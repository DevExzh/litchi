//! Create a Pages document and native chart without an input package.

use std::env;

use litchi_iwa::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartAxisTickMarkLocation, ChartCornerRadius, ChartData, ChartKind, ChartRoundedCorners,
    ChartValueAxisBounds, ChartValueAxisScale, ChartValueAxisSteps,
};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_pages_chart <output.pages>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let body = "Quarterly Results";
    let mut editor = PagesDocumentBuilder::new().body_text(body).build()?;
    let data = ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?;
    let chart = editor.add_body_chart(
        body.encode_utf16().count(),
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 96.0, y: 144.0 },
        DrawableSize {
            width: 360.0,
            height: 240.0,
        },
    )?;
    editor.set_body_chart_title(chart.drawable_object_id, "Quarterly revenue")?;
    editor.set_body_chart_border_visible(chart.drawable_object_id, true)?;
    editor.set_body_chart_rounded_corners(
        chart.drawable_object_id,
        ChartRoundedCorners::new(ChartCornerRadius::new(20.0)?, true),
    )?;
    editor.set_body_chart_axis_title(chart.drawable_object_id, ChartAxis::Category, "Quarter")?;
    editor.set_body_chart_axis_title(chart.drawable_object_id, ChartAxis::Value, "Revenue")?;
    editor.set_body_chart_value_axis_bounds(
        chart.drawable_object_id,
        ChartValueAxisBounds::fixed(ChartAxisBound::new(1.0)?, ChartAxisBound::new(30.0)?)?,
    )?;
    editor.set_body_chart_value_axis_scale(
        chart.drawable_object_id,
        ChartValueAxisScale::Logarithmic,
    )?;
    editor.set_body_chart_value_axis_steps(
        chart.drawable_object_id,
        ChartValueAxisSteps::fixed(
            ChartAxisMajorStepCount::new(6)?,
            ChartAxisMinorStepCount::new(2)?,
        ),
    )?;
    editor.set_body_chart_value_axis_minimum_label_visible(chart.drawable_object_id, false)?;
    editor.set_body_chart_category_axis_series_names_visible(chart.drawable_object_id, true)?;
    editor.set_body_chart_axis_labels_visible(
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_body_chart_axis_minor_tick_marks_visible(
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_body_chart_axis_tick_mark_location(
        chart.drawable_object_id,
        ChartAxis::Category,
        ChartAxisTickMarkLocation::Outside,
    )?;
    editor.set_body_chart_axis_line_visible(chart.drawable_object_id, ChartAxis::Value, false)?;
    editor.set_body_chart_axis_major_gridlines_visible(
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_body_chart_axis_minor_gridlines_visible(
        chart.drawable_object_id,
        ChartAxis::Value,
        true,
    )?;
    editor.set_body_chart_legend_visible(chart.drawable_object_id, false)?;
    editor.set_body_chart_caption(chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Pages {:?} chart {} with native chart and axis titles, a visible chart border, 20% rounded outside corners, a logarithmic value-axis scale with fixed bounds and steps, hidden category-axis labels and minor tick marks, outside category-axis major tick marks, a hidden value-axis minimum label, line, and legend, visible category-axis series names, hidden value-axis major gridlines, visible value-axis minor gridlines, and a caption at body UTF-16 index {}",
        chart.kind, chart.drawable_object_id, chart.anchor_character_index
    );
    Ok(())
}
