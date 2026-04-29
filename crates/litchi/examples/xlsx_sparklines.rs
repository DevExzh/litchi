use litchi::ooxml::xlsx::{
    Sparkline, SparklineAxisMinMax, SparklineColor, SparklineDisplayEmptyCellsAs, SparklineGroup,
    SparklineType, Workbook,
};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let mut wb = Workbook::create()?;
    let sheet = wb.worksheet_mut(0)?;
    sheet.set_name("Sparklines".to_string());

    sheet.set_cell_value(1, 1, "Sparkline");
    for i in 0..10u32 {
        sheet.set_cell_value(1, i + 2, (i + 1) as f64);
    }

    for row in 0..3u32 {
        let r = row + 2;
        for i in 0..10u32 {
            let v = ((row + 1) * (i + 2)) as f64;
            sheet.set_cell_value(r, i + 2, v);
        }
    }

    let mut group = SparklineGroup::new(SparklineType::Column);
    group.options.display_empty_cells_as = SparklineDisplayEmptyCellsAs::Gap;
    group.options.display_x_axis = true;
    group.options.markers = true;
    group.options.high = true;
    group.options.low = true;
    group.options.first = true;
    group.options.last = true;
    group.options.negative = true;
    group.options.min_axis_type = SparklineAxisMinMax::Custom;
    group.options.max_axis_type = SparklineAxisMinMax::Custom;
    group.options.manual_min = Some(0.0);
    group.options.manual_max = Some(40.0);
    group.options.line_weight = Some(0.75);

    group.colors.series = Some(SparklineColor::new("FF1F77B4".to_string()));
    group.colors.axis = Some(SparklineColor::new("FF333333".to_string()));
    group.colors.markers = Some(SparklineColor::new("FFFF7F0E".to_string()));
    group.colors.high = Some(SparklineColor::new("FF2CA02C".to_string()));
    group.colors.low = Some(SparklineColor::new("FFD62728".to_string()));
    group.colors.first = Some(SparklineColor::new("FF9467BD".to_string()));
    group.colors.last = Some(SparklineColor::new("FF8C564B".to_string()));
    group.colors.negative = Some(SparklineColor::new("FFBCBD22".to_string()));

    for row in 0..3u32 {
        let r = row + 2;
        group.push(Sparkline {
            data_range: format!("Sparklines!B{r}:K{r}"),
            location: format!("A{r}"),
        });
    }
    sheet.add_sparkline_group(group);

    let args: Vec<String> = env::args().collect();
    let out = args
        .get(1)
        .map(String::as_str)
        .unwrap_or("sparklines_output.xlsx");
    wb.save(out)?;
    println!("Wrote {}", out);
    Ok(())
}
