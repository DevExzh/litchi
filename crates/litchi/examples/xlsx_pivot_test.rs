//! Typed SpreadsheetML pivot-table model showcase.
//!
//! The XLSX facade currently exposes pivot authoring at the typed cache/table
//! part boundary. This example therefore writes transactional source
//! workbooks alongside the cache-definition, cache-record, and pivot-table
//! XML models, then reads each model back through the same facade.
//!
//! ```bash
//! cargo run -p litchi --example xlsx_pivot_test --features ooxml -- pivot_examples
//! # -> pivot_examples/pivot_basic.xlsx
//! # -> pivot_examples/pivot_basic.pivot-cache.xml
//! # -> pivot_examples/pivot_basic.pivot-records.xml
//! # -> pivot_examples/pivot_basic.pivot-table.xml
//! ```

use litchi_xlsx::Workbook;
use litchi_xlsx::pivot::cache::{CacheRecord, Item, Records};
use litchi_xlsx::pivot::cache::{Definition as CacheDefinition, Field as CacheField};
use litchi_xlsx::pivot::fields::{AxisField, AxisItem, DataField, Field, PageField, Subtotal};
use litchi_xlsx::pivot::filters::Filter;
use litchi_xlsx::pivot::styles::{Location, Style};
use litchi_xlsx::pivot::writer::TableDefinition;
use litchi_xlsx::pivot::{
    AxisType, ItemType, PivotTable, SortType, read_pivot_cache_definition,
    read_pivot_cache_records, read_pivot_table_definition, write_pivot_cache_definition,
    write_pivot_cache_records, write_pivot_table,
};
use std::env;
use std::fs;
use std::path::Path;

type ExampleResult<T> = Result<T, Box<dyn std::error::Error + Send + Sync>>;

fn main() -> ExampleResult<()> {
    let output_dir = env::args()
        .nth(1)
        .unwrap_or_else(|| "pivot_examples".to_string());
    fs::create_dir_all(&output_dir)?;

    generate_basic_example(Path::new(&output_dir))?;
    generate_filter_example(Path::new(&output_dir))?;

    println!("Generated typed pivot artifacts in {}", output_dir);
    Ok(())
}

fn generate_basic_example(output_dir: &Path) -> ExampleResult<()> {
    let workbook_path = output_dir.join("pivot_basic.xlsx");
    let model_path = output_dir.join("pivot_basic.pivot-table.xml");
    let cache_path = output_dir.join("pivot_basic.pivot-cache.xml");
    let records_path = output_dir.join("pivot_basic.pivot-records.xml");

    let records = [
        ["Laptops", "North", "Q1", "120000"],
        ["Laptops", "North", "Q2", "98500"],
        ["Laptops", "South", "Q1", "76200"],
        ["Laptops", "South", "Q2", "88450"],
        ["Tablets", "North", "Q1", "54300"],
        ["Tablets", "North", "Q2", "61125"],
        ["Tablets", "South", "Q1", "42000"],
        ["Tablets", "South", "Q2", "39875"],
        ["Accessories", "North", "Q1", "15200"],
        ["Accessories", "North", "Q2", "12950"],
        ["Accessories", "South", "Q1", "10400"],
        ["Accessories", "South", "Q2", "9875"],
    ];
    let workbook = source_workbook(
        "SalesData",
        ["Product", "Region", "Quarter", "Amount"],
        &records,
        "PivotSummary",
        "Sales by product and quarter",
    )?;
    workbook.save(&workbook_path)?;

    let cache = cache_definition(
        "SalesData",
        "A1:D13",
        [
            ("Product", &["Laptops", "Tablets", "Accessories"][..]),
            ("Region", &["North", "South"][..]),
            ("Quarter", &["Q1", "Q2"][..]),
            ("Amount", &[][..]),
        ],
        records.len(),
    );
    let cache_records = records_model(records.iter().map(|row| {
        row.iter().enumerate().map(|(index, value)| {
            if index == 3 {
                Item::Number(value.parse().expect("static amount"))
            } else {
                Item::String((*value).to_string())
            }
        })
    }));
    let fields = pivot_fields(
        ["Product", "Region", "Quarter", "Amount"],
        [
            Some(AxisType::AxisRow),
            None,
            Some(AxisType::AxisCol),
            Some(AxisType::AxisValues),
        ],
        [false, false, false, true],
    );
    let table = TableDefinition {
        name: "SalesByProductQuarter",
        cache_id: 0,
        location: &Location {
            reference: "A3".into(),
            first_header_row: 1,
            first_data_row: 2,
            first_data_col: 1,
            ..Location::default()
        },
        pivot_fields: &fields,
        row_fields: &[AxisField { x: 0 }],
        col_fields: &[AxisField { x: 2 }],
        page_fields: &[],
        data_fields: &[DataField {
            name: Some("Total Amount".into()),
            fld: 3,
            subtotal: Subtotal::Sum,
            ..DataField::default()
        }],
        row_items: &[AxisItem {
            item_type: ItemType::Data,
            ..AxisItem::default()
        }],
        col_items: &[],
        filters: &[],
        style: Some(&Style {
            name: Some("PivotStyleMedium2".into()),
            show_row_headers: Some(true),
            show_col_headers: Some(true),
            ..Style::default()
        }),
    };
    write_models(
        &cache,
        &cache_records,
        &table,
        &cache_path,
        &records_path,
        &model_path,
        "SalesByProductQuarter",
        &["Product", "Region", "Quarter", "Amount"],
    )?;
    Ok(())
}

fn generate_filter_example(output_dir: &Path) -> ExampleResult<()> {
    let workbook_path = output_dir.join("pivot_with_filters.xlsx");
    let model_path = output_dir.join("pivot_with_filters.pivot-table.xml");
    let cache_path = output_dir.join("pivot_with_filters.pivot-cache.xml");
    let records_path = output_dir.join("pivot_with_filters.pivot-records.xml");

    let records = [
        ["Alice", "Online", "East", "2024", "45000"],
        ["Alice", "Retail", "East", "2025", "52000"],
        ["Bob", "Online", "West", "2024", "38000"],
        ["Bob", "Retail", "West", "2025", "41500"],
        ["Carol", "Online", "East", "2024", "33750"],
        ["Carol", "Retail", "East", "2025", "36200"],
        ["Dave", "Online", "West", "2024", "29600"],
        ["Dave", "Retail", "West", "2025", "31400"],
    ];
    let workbook = source_workbook(
        "Opportunities",
        ["Salesperson", "Channel", "Region", "Year", "Revenue"],
        &records,
        "PivotWithFilters",
        "Revenue by channel",
    )?;
    workbook.save(&workbook_path)?;

    let cache = cache_definition(
        "Opportunities",
        "A1:E9",
        [
            ("Salesperson", &["Alice", "Bob", "Carol", "Dave"][..]),
            ("Channel", &["Online", "Retail"][..]),
            ("Region", &["East", "West"][..]),
            ("Year", &["2024", "2025"][..]),
            ("Revenue", &[][..]),
        ],
        records.len(),
    );
    let cache_records = records_model(records.iter().map(|row| {
        row.iter().enumerate().map(|(index, value)| {
            if index >= 3 {
                Item::Number(value.parse().expect("static numeric value"))
            } else {
                Item::String((*value).to_string())
            }
        })
    }));
    let fields = pivot_fields(
        ["Salesperson", "Channel", "Region", "Year", "Revenue"],
        [
            None,
            Some(AxisType::AxisCol),
            Some(AxisType::AxisRow),
            Some(AxisType::AxisPage),
            Some(AxisType::AxisValues),
        ],
        [false, false, false, false, true],
    );
    let table = TableDefinition {
        name: "RevenueByChannel",
        cache_id: 0,
        location: &Location {
            reference: "A3".into(),
            first_header_row: 1,
            first_data_row: 2,
            first_data_col: 1,
            ..Location::default()
        },
        pivot_fields: &fields,
        row_fields: &[AxisField { x: 2 }],
        col_fields: &[AxisField { x: 1 }],
        page_fields: &[PageField {
            fld: 3,
            item: None,
            hier: None,
            name: Some("Year".into()),
            cap: None,
        }],
        data_fields: &[DataField {
            name: Some("Revenue".into()),
            fld: 4,
            subtotal: Subtotal::Sum,
            ..DataField::default()
        }],
        row_items: &[AxisItem::default()],
        col_items: &[],
        filters: &[Filter {
            fld: 3,
            filter_type: "captionEqual".into(),
            id: 0,
            string_value1: Some("2025".into()),
            ..Filter::default()
        }],
        style: Some(&Style {
            name: Some("PivotStyleMedium9".into()),
            show_row_headers: Some(true),
            show_col_headers: Some(true),
            show_row_stripes: Some(true),
            ..Style::default()
        }),
    };
    write_models(
        &cache,
        &cache_records,
        &table,
        &cache_path,
        &records_path,
        &model_path,
        "RevenueByChannel",
        &["Salesperson", "Channel", "Region", "Year", "Revenue"],
    )?;
    Ok(())
}

fn source_workbook<const C: usize>(
    source_name: &str,
    headers: [&str; C],
    rows: &[[&str; C]],
    summary_name: &str,
    summary_caption: &str,
) -> ExampleResult<Workbook> {
    let workbook = Workbook::create()?;
    let mut edit = workbook.edit()?;
    edit.tab(0)?
        .ok_or("missing initial worksheet")?
        .rename(source_name)?;
    {
        let mut sheet = edit.sheet(source_name)?.ok_or("missing source worksheet")?;
        for (column, header) in headers.iter().enumerate() {
            sheet.set((1_u32, (column + 1) as u32), *header)?;
        }
        for (row, values) in rows.iter().enumerate() {
            for (column, value) in values.iter().enumerate() {
                sheet.set(((row + 2) as u32, (column + 1) as u32), *value)?;
            }
        }
    }
    let mut summary = edit.add(summary_name)?;
    summary.set("A1", summary_caption)?;
    Ok(edit.commit()?.into_workbook())
}

fn cache_definition<const N: usize>(
    source_sheet: &str,
    source_ref: &str,
    fields: [(&str, &[&str]); N],
    record_count: usize,
) -> CacheDefinition {
    CacheDefinition {
        id: Some("rIdPivotCacheRecords".into()),
        source_worksheet: Some(source_sheet.into()),
        source_ref: Some(source_ref.into()),
        record_count: Some(record_count as u32),
        cache_fields: fields
            .into_iter()
            .map(|(name, values)| CacheField {
                name: name.into(),
                shared_items: values
                    .iter()
                    .map(|value| litchi_xlsx::pivot::cache::Item::String((*value).into()))
                    .collect(),
                ..CacheField::default()
            })
            .collect(),
        ..CacheDefinition::default()
    }
}

fn records_model<I, J>(rows: I) -> Records
where
    I: IntoIterator<Item = J>,
    J: IntoIterator<Item = Item>,
{
    Records {
        records: rows
            .into_iter()
            .map(|values| CacheRecord {
                values: values.into_iter().collect(),
            })
            .collect(),
    }
}

fn pivot_fields<const N: usize>(
    names: [&str; N],
    axes: [Option<AxisType>; N],
    data_fields: [bool; N],
) -> Vec<Field> {
    names
        .into_iter()
        .zip(axes)
        .zip(data_fields)
        .map(|((name, axis), data_field)| Field {
            name: Some(name.into()),
            axis,
            data_field: Some(data_field),
            sort_type: SortType::Manual,
            ..Field::default()
        })
        .collect()
}

fn write_models(
    cache: &CacheDefinition,
    records: &Records,
    table: &TableDefinition<'_>,
    cache_path: &Path,
    records_path: &Path,
    table_path: &Path,
    expected_name: &str,
    expected_fields: &[&str],
) -> ExampleResult<()> {
    let cache_xml = write_pivot_cache_definition(cache)?;
    let records_xml = write_pivot_cache_records(records)?;
    let table_xml = write_pivot_table(table)?;

    let parsed_cache = read_pivot_cache_definition(&cache_xml)?.ok_or("missing cache root")?;
    let parsed_records = read_pivot_cache_records(&records_xml)?.ok_or("missing records root")?;
    let parsed_table = read_pivot_table_definition(&table_xml)?.ok_or("missing table root")?;
    verify_model(
        &parsed_cache,
        &parsed_records,
        &parsed_table,
        expected_name,
        expected_fields,
        records.records.len(),
    )?;

    fs::write(cache_path, cache_xml)?;
    fs::write(records_path, records_xml)?;
    fs::write(table_path, table_xml)?;
    Ok(())
}

fn verify_model(
    cache: &CacheDefinition,
    records: &Records,
    table: &PivotTable,
    expected_name: &str,
    expected_fields: &[&str],
    expected_records: usize,
) -> ExampleResult<()> {
    if cache
        .cache_fields
        .iter()
        .map(|field| field.name.as_str())
        .collect::<Vec<_>>()
        != expected_fields
    {
        return Err("pivot cache field names did not survive the codec round trip".into());
    }
    if records.records.len() != expected_records {
        return Err("pivot cache record count did not survive the codec round trip".into());
    }
    let expected_fields = expected_fields
        .iter()
        .map(|name| (*name).to_string())
        .collect::<Vec<_>>();
    if table.name != expected_name || table.field_names != expected_fields {
        return Err("pivot table identity did not survive the codec round trip".into());
    }
    Ok(())
}
