//! Inspect native chart metadata in a Pages, Numbers, or Keynote package.

use std::env;

use litchi_iwa::Document;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let input = arguments
        .next()
        .ok_or("usage: inspect_iwork_charts <input>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }

    let document = Document::open(input)?;
    let charts = document.charts()?;
    for chart in &charts {
        println!(
            "object={} kind={:?} series={} default_data={} title={:?}",
            chart.object_id,
            chart.chart_type,
            chart.series_count,
            chart.contains_default_data,
            chart.title
        );
        println!("  rows={:?}", chart.row_names);
        println!("  columns={:?}", chart.column_names);
    }
    println!("charts={}", charts.len());
    Ok(())
}
