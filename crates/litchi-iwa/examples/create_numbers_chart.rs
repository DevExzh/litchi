//! Create a Numbers spreadsheet and standalone chart without an input package.

use std::env;

use litchi_iwa::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartAxisTickMarkLocation, ChartCornerRadius, ChartData, ChartGapPercentage, ChartGapSpacing,
    ChartKind, ChartRoundedCorners, ChartSeriesValueLabelVisibility, ChartShadow,
    ChartValueAxisBounds, ChartValueAxisScale, ChartValueAxisSteps,
};
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeDropShadow, ShapeFill,
    ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeStroke, StrokePattern, StrokeWidth,
};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_numbers_chart <output.numbers>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let mut editor = NumbersDocumentBuilder::new()
        .sheet_name("Quarterly Results")
        .table_name("Source Data")
        .build()?;
    let sheet_id = editor.sheets()?[0].object_id;
    let data = ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?;
    let chart = editor.add_sheet_chart(
        sheet_id,
        ChartKind::Column2d,
        data,
        DrawablePoint { x: 420.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    editor.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Quarterly revenue")?;
    editor.set_sheet_chart_background_fill(
        sheet_id,
        chart.drawable_object_id,
        &ShapeFill::Solid(RgbaColor::new(0.85, 0.92, 1.0, 1.0, RgbColorSpace::Srgb)?),
    )?;
    editor.set_sheet_chart_border_visible(sheet_id, chart.drawable_object_id, true)?;
    editor.set_sheet_chart_border_stroke(
        sheet_id,
        chart.drawable_object_id,
        Some(ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb)?,
            StrokeWidth::new(3.0)?,
            StrokePattern::MediumDash,
        )),
    )?;
    editor.set_sheet_chart_rounded_corners(
        sheet_id,
        chart.drawable_object_id,
        ChartRoundedCorners::new(ChartCornerRadius::new(20.0)?, true),
    )?;
    editor.set_sheet_chart_gap_spacing(
        sheet_id,
        chart.drawable_object_id,
        ChartGapSpacing::new(
            ChartGapPercentage::new(25.0)?,
            ChartGapPercentage::new(70.0)?,
        ),
    )?;
    editor.set_sheet_chart_shadow(
        sheet_id,
        chart.drawable_object_id,
        ChartShadow::Grouped(ShapeDropShadow::new(
            ShapeShadowAppearance::new(
                RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb)?,
                ShapeShadowBlurRadius::from_points(15)?,
                ShapeShadowOffset::from_points(8.0)?,
                ShapeShadowOpacity::new(0.6)?,
            ),
            ShapeShadowAngle::from_degrees(60.0)?,
        )),
    )?;
    editor.set_sheet_chart_axis_title(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Category,
        "Quarter",
    )?;
    editor.set_sheet_chart_axis_title(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        "Revenue",
    )?;
    editor.set_sheet_chart_value_axis_bounds(
        sheet_id,
        chart.drawable_object_id,
        ChartValueAxisBounds::fixed(ChartAxisBound::new(1.0)?, ChartAxisBound::new(30.0)?)?,
    )?;
    editor.set_sheet_chart_value_axis_scale(
        sheet_id,
        chart.drawable_object_id,
        ChartValueAxisScale::Logarithmic,
    )?;
    editor.set_sheet_chart_value_axis_steps(
        sheet_id,
        chart.drawable_object_id,
        ChartValueAxisSteps::fixed(
            ChartAxisMajorStepCount::new(6)?,
            ChartAxisMinorStepCount::new(2)?,
        ),
    )?;
    editor.set_sheet_chart_value_axis_minimum_label_visible(
        sheet_id,
        chart.drawable_object_id,
        false,
    )?;
    editor.set_sheet_chart_category_axis_series_names_visible(
        sheet_id,
        chart.drawable_object_id,
        true,
    )?;
    editor.set_sheet_chart_axis_labels_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_sheet_chart_axis_minor_tick_marks_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Category,
        false,
    )?;
    editor.set_sheet_chart_axis_tick_mark_location(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Category,
        ChartAxisTickMarkLocation::Outside,
    )?;
    editor.set_sheet_chart_axis_line_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_sheet_chart_axis_major_gridlines_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        false,
    )?;
    editor.set_sheet_chart_axis_minor_gridlines_visible(
        sheet_id,
        chart.drawable_object_id,
        ChartAxis::Value,
        true,
    )?;
    editor.set_sheet_chart_includes_hidden_data(sheet_id, chart.drawable_object_id, false)?;
    editor.set_sheet_chart_legend_visible(sheet_id, chart.drawable_object_id, false)?;
    editor.set_sheet_chart_series_value_label_visibilities(
        sheet_id,
        chart.drawable_object_id,
        &[ChartSeriesValueLabelVisibility::Visible; 2],
    )?;
    editor.set_sheet_chart_caption(sheet_id, chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Numbers {:?} chart {} with native chart and axis titles, a light-blue color background, a visible blue 3 pt medium-dash chart border, a grouped blue 15 pt shadow, 20% rounded outside corners, 25% item and 70% set gaps, a logarithmic value-axis scale with fixed bounds and steps, hidden category-axis labels and minor tick marks, outside category-axis major tick marks, a hidden value-axis minimum label, line, and legend, visible category-axis series names and data value labels, hidden value-axis major gridlines, visible value-axis minor gridlines, excluded hidden source rows and columns, and a caption on sheet {}",
        chart.kind, chart.drawable_object_id, sheet_id
    );
    Ok(())
}
