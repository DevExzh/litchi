//! Create scratch Pages, Numbers, and Keynote charts with custom category labels.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{ChartData, Kind};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize};
use litchi_iwa_common::chart::category_labels::{Frequency, Interval, Layout};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_category_label_layouts <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;
    let layout = Layout::new(Frequency::Every(Interval::new(3)?), false);

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Category Label Layout")
        .build()?;
    let sheet_id = numbers.sheets()?[0].id();
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 500.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Every third category")?;
    numbers.set_sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id, layout)?;
    assert_eq!(
        numbers.sheet_chart_category_label_layout(sheet_id, chart.drawable_object_id)?,
        layout
    );
    numbers.save(output.join("category-label-layout-crate.numbers"))?;

    let body = "Category label layout CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 500.0,
            height: 300.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Every third category")?;
    pages.set_body_chart_category_label_layout(chart.drawable_object_id, layout)?;
    assert_eq!(
        pages.body_chart_category_label_layout(chart.drawable_object_id)?,
        layout
    );
    pages.save(output.join("category-label-layout-crate.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Category Label Layout CRUD")
        .build()?;
    let chart = keynote.add_slide_chart(
        0,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 260.0, y: 220.0 },
        DrawableSize {
            width: 1_400.0,
            height: 650.0,
        },
    )?;
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Every third category")?;
    keynote.set_slide_chart_category_label_layout(0, chart.drawable_object_id, layout)?;
    assert_eq!(
        keynote.slide_chart_category_label_layout(0, chart.drawable_object_id)?,
        layout
    );
    keynote.save(output.join("category-label-layout-crate.key"))?;

    println!(
        "created category-label-layout fixtures in {}",
        output.display()
    );
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    let categories = (1..=12).map(|month| format!("M{month}")).collect();
    Ok(ChartData::new(
        vec!["Revenue".to_owned(), "Cost".to_owned()],
        categories,
        vec![
            vec![
                Some(12.0),
                Some(18.0),
                Some(15.0),
                Some(21.0),
                Some(25.0),
                Some(22.0),
                Some(28.0),
                Some(31.0),
                Some(29.0),
                Some(35.0),
                Some(38.0),
                Some(42.0),
            ],
            vec![
                Some(8.0),
                Some(10.0),
                Some(9.0),
                Some(13.0),
                Some(14.0),
                Some(15.0),
                Some(17.0),
                Some(19.0),
                Some(18.0),
                Some(21.0),
                Some(23.0),
                Some(25.0),
            ],
        ],
    )?)
}
