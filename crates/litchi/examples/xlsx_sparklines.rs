//! Typed sparkline model with transactional XLSX source data.
//!
//! Sparkline extension attachment is not part of the ordinary worksheet
//! transaction yet. The typed group is validated by its XML codec while the
//! source values are committed through the standalone workbook facade.

use litchi_xlsx::sheet::sparklines::{
    AxisMinMax, Color, DisplayEmptyCellsAs, Group, Item, Type, write_groups_ext,
};
use litchi_xlsx::{Number, Workbook};
use std::env;

fn main() -> Result<(), Box<dyn std::error::Error + Send + Sync>> {
    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("default worksheet is missing")?
        .rename("Sparklines")?;
    {
        let mut sheet = edit
            .sheet("Sparklines")?
            .ok_or("Sparklines worksheet is missing")?;
        sheet.set("A1", "Sparkline")?;
        for column in 0..10 {
            sheet.set((1, column + 1), Number::new((column + 1).to_string())?)?;
        }
        for row in 0..3 {
            let spreadsheet_row = row + 2;
            for column in 0..10 {
                sheet.set(
                    (spreadsheet_row, column + 1),
                    Number::new(((row + 1) * (column + 2)).to_string())?,
                )?;
            }
        }
    }

    let mut group = Group::new(Type::Column);
    group.options.display_empty_cells_as = DisplayEmptyCellsAs::Gap;
    group.options.display_x_axis = true;
    group.options.markers = true;
    group.options.high = true;
    group.options.low = true;
    group.options.first = true;
    group.options.last = true;
    group.options.negative = true;
    group.options.min_axis_type = AxisMinMax::Custom;
    group.options.max_axis_type = AxisMinMax::Custom;
    group.options.manual_min = Some(0.0);
    group.options.manual_max = Some(40.0);
    group.options.line_weight = Some(0.75);
    group.colors.series = Some(Color::new("FF1F77B4".to_string()));
    group.colors.axis = Some(Color::new("FF333333".to_string()));
    group.colors.markers = Some(Color::new("FFFF7F0E".to_string()));
    group.colors.high = Some(Color::new("FF2CA02C".to_string()));
    group.colors.low = Some(Color::new("FFD62728".to_string()));
    group.colors.first = Some(Color::new("FF9467BD".to_string()));
    group.colors.last = Some(Color::new("FF8C564B".to_string()));
    group.colors.negative = Some(Color::new("FFBCBD22".to_string()));
    for row in 2..=4 {
        group.push(Item {
            data_range: format!("Sparklines!B{row}:K{row}"),
            location: format!("A{row}"),
        });
    }
    let mut extension_xml = String::new();
    write_groups_ext(&mut extension_xml, &[group])?;

    let output = env::args()
        .nth(1)
        .unwrap_or_else(|| "sparklines_output.xlsx".to_string());
    edit.commit()?.into_workbook().save(&output)?;
    println!(
        "Wrote {output}; validated sparkline extension ({} bytes)",
        extension_xml.len()
    );
    Ok(())
}
