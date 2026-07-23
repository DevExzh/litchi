//! Create a Keynote presentation and native chart without an input package.

use std::env;

use litchi_iwa::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartAxisTickMarkLocation, ChartData, ChartKind, ChartValueAxisBounds, ChartValueAxisScale,
    ChartValueAxisSteps,
};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_keynote_chart <output.key>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = KeynoteDocumentBuilder::new()
        .title("Quarterly Results")
        .subtitle("Native chart built from typed IWA objects")
        .build()?;
    let data = ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?;
    let chart = editor.add_slide_chart(
        0,
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 240.0, y: 300.0 },
        DrawableSize {
            width: 1_440.0,
            height: 600.0,
        },
    )?;
    editor.set_slide_chart_title(0, chart.drawable_object_id, "Quarterly revenue")?;
    editor.set_slide_chart_border_visible(0, chart.drawable_object_id, true)?;
    editor.set_slide_chart_axis_title(
        0,
        chart.drawable_object_id,
        ChartAxis::Category,
        "Quarter",
    )?;
    editor.set_slide_chart_axis_title(0, chart.drawable_object_id, ChartAxis::Value, "Revenue")?;
    editor.set_slide_chart_value_axis_bounds(
        0,
        chart.drawable_object_id,
        ChartValueAxisBounds::fixed(ChartAxisBound::new(1.0)?, ChartAxisBound::new(30.0)?)?,
    )?;
    editor.set_slide_chart_value_axis_scale(
        0,
        chart.drawable_object_id,
        ChartValueAxisScale::Logarithmic,
    )?;
    editor.set_slide_chart_value_axis_steps(
        0,
        chart.drawable_object_id,
        ChartValueAxisSteps::fixed(
            ChartAxisMajorStepCount::new(6)?,
            ChartAxisMinorStepCount::new(2)?,
        ),
    )?;
    editor.set_slide_chart_value_axis_minimum_label_visible(0, chart.drawable_object_id, false)?;
    editor.set_slide_chart_category_axis_series_names_visible(0, chart.drawable_object_id, true)?;
    editor.set_slide_chart_axis_labels_visible(
        0,
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_slide_chart_axis_minor_tick_marks_visible(
        0,
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_slide_chart_axis_tick_mark_location(
        0,
        chart.drawable_object_id,
        ChartAxis::Category,
        ChartAxisTickMarkLocation::Outside,
    )?;
    editor.set_slide_chart_axis_line_visible(
        0,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_slide_chart_axis_major_gridlines_visible(
        0,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_slide_chart_axis_minor_gridlines_visible(
        0,
        chart.drawable_object_id,
        ChartAxis::Value,
        true,
    )?;
    editor.set_slide_chart_legend_visible(0, chart.drawable_object_id, false)?;
    editor.set_slide_chart_caption(0, chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Keynote {:?} chart {} with native chart and axis titles, a visible chart border, a logarithmic value-axis scale with fixed bounds and steps, hidden category-axis labels and minor tick marks, outside category-axis major tick marks, a hidden value-axis minimum label, line, and legend, visible category-axis series names, hidden value-axis major gridlines, visible value-axis minor gridlines, and a caption on slide {}",
        chart.kind, chart.drawable_object_id, chart.slide_index
    );
    Ok(())
}
