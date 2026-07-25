//! Create a Pages document and native chart without an input package.

use std::env;

use litchi_iwa::charts::{
    ChartAxis, ChartAxisBound, ChartAxisMajorStepCount, ChartAxisMinorStepCount,
    ChartAxisTickMarkLocation, ChartCornerRadius, ChartData, ChartErrorBarDirection,
    ChartErrorBarFixedValue, ChartErrorBarPercentage, ChartGapPercentage, ChartGapSpacing,
    ChartKind, ChartRoundedCorners, ChartSeriesErrorBarAutoFit, ChartSeriesErrorBars,
    ChartSeriesTrendline, ChartSeriesTrendlineMovingAveragePeriod, ChartSeriesValueLabelAffixes,
    ChartSeriesValueLabelAutoFit, ChartSeriesValueLabelDecimalPlaces,
    ChartSeriesValueLabelLocation, ChartSeriesValueLabelNegativeStyle,
    ChartSeriesValueLabelNumberFormat, ChartSeriesValueLabelVisibility, ChartShadow,
    ChartValueAxisBounds, ChartValueAxisScale, ChartValueAxisSteps,
};
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{
    DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeDropShadow, ShapeFill,
    ShapeShadowAngle, ShapeShadowAppearance, ShapeShadowBlurRadius, ShapeShadowOffset,
    ShapeShadowOpacity, ShapeStroke, StrokePattern, StrokeWidth,
};

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
    editor.set_body_chart_background_fill(
        chart.drawable_object_id,
        &ShapeFill::Solid(RgbaColor::new(0.85, 0.92, 1.0, 1.0, RgbColorSpace::Srgb)?),
    )?;
    editor.set_body_chart_border_visible(chart.drawable_object_id, true)?;
    editor.set_body_chart_border_stroke(
        chart.drawable_object_id,
        Some(ShapeStroke::new(
            RgbaColor::new(0.1, 0.3, 0.8, 1.0, RgbColorSpace::Srgb)?,
            StrokeWidth::new(3.0)?,
            StrokePattern::MediumDash,
        )),
    )?;
    editor.set_body_chart_rounded_corners(
        chart.drawable_object_id,
        ChartRoundedCorners::new(ChartCornerRadius::new(20.0)?, true),
    )?;
    editor.set_body_chart_gap_spacing(
        chart.drawable_object_id,
        ChartGapSpacing::new(
            ChartGapPercentage::new(25.0)?,
            ChartGapPercentage::new(70.0)?,
        ),
    )?;
    editor.set_body_chart_shadow(
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
    editor.set_body_chart_series_value_label_visibilities(
        chart.drawable_object_id,
        &[ChartSeriesValueLabelVisibility::Visible; 2],
    )?;
    editor.set_body_chart_series_value_label_locations(
        chart.drawable_object_id,
        &[ChartSeriesValueLabelLocation::Outside; 2],
    )?;
    editor.set_body_chart_series_value_label_affixes(
        chart.drawable_object_id,
        &[
            ChartSeriesValueLabelAffixes::new("$", " USD"),
            ChartSeriesValueLabelAffixes::new("€", " EUR"),
        ],
    )?;
    editor.set_body_chart_series_value_label_number_formats(
        chart.drawable_object_id,
        &[
            ChartSeriesValueLabelNumberFormat::new(
                ChartSeriesValueLabelDecimalPlaces::fixed(2)?,
                ChartSeriesValueLabelNegativeStyle::Parentheses,
                false,
            ),
            ChartSeriesValueLabelNumberFormat::new(
                ChartSeriesValueLabelDecimalPlaces::fixed(1)?,
                ChartSeriesValueLabelNegativeStyle::MinusSign,
                true,
            ),
        ],
    )?;
    editor.set_body_chart_series_value_label_auto_fits(
        chart.drawable_object_id,
        &[
            ChartSeriesValueLabelAutoFit::Disabled,
            ChartSeriesValueLabelAutoFit::Enabled,
        ],
    )?;
    editor.set_body_chart_series_trendlines(
        chart.drawable_object_id,
        &[
            ChartSeriesTrendline::linear()
                .with_legend_name("Revenue fit")?
                .with_equation_visibility(true)?
                .with_r_squared_visibility(true)?,
            ChartSeriesTrendline::moving_average(ChartSeriesTrendlineMovingAveragePeriod::new(2)?)
                .with_legend_visibility(true)?,
        ],
    )?;
    editor.set_body_chart_series_error_bars(
        chart.drawable_object_id,
        &[
            ChartSeriesErrorBars::FixedValue {
                direction: ChartErrorBarDirection::PositiveAndNegative,
                value: ChartErrorBarFixedValue::new(12.5)?,
            },
            ChartSeriesErrorBars::Percentage {
                direction: ChartErrorBarDirection::PositiveOnly,
                percentage: ChartErrorBarPercentage::new(17)?,
            },
        ],
    )?;
    editor.set_body_chart_series_error_bar_auto_fits(
        chart.drawable_object_id,
        &[
            ChartSeriesErrorBarAutoFit::Disabled,
            ChartSeriesErrorBarAutoFit::Enabled,
        ],
    )?;
    editor.set_body_chart_caption(chart.drawable_object_id, "Revenue by region")?;
    editor.save(output)?;
    println!(
        "created Pages {:?} chart {} with native chart and axis titles, a light-blue color background, a visible blue 3 pt medium-dash chart border, a grouped blue 15 pt shadow, 20% rounded outside corners, 25% item and 70% set gaps, a logarithmic value-axis scale with fixed bounds and steps, hidden category-axis labels and minor tick marks, outside category-axis major tick marks, a hidden value-axis minimum label, line, and legend, visible category-axis series names, explicitly number-formatted currency-affixed data value labels placed outside with per-series Auto-Fit, linear and moving-average series trendlines, fixed and percentage series error bars with per-series Auto-Fit, hidden value-axis major gridlines, visible value-axis minor gridlines, and a caption at body UTF-16 index {}",
        chart.kind, chart.drawable_object_id, chart.anchor_character_index
    );
    Ok(())
}
