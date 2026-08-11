//! Create Pages, Numbers, and Keynote line charts with typed data symbols.

use std::env;
use std::path::Path;

use litchi_iwa::charts::{
    ChartData, ChartSeriesConnectionLine, ChartSeriesStroke, ChartSeriesStrokePattern,
    ChartSeriesSymbol, ChartSeriesSymbolFill, ChartSeriesSymbolShape, ChartSeriesSymbolSize, Kind,
};
use litchi_iwa::keynote::KeynoteDocumentBuilder;
use litchi_iwa::numbers::NumbersDocumentBuilder;
use litchi_iwa::pages::PagesDocumentBuilder;
use litchi_iwa::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor, ShapeFill, Width};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut arguments = env::args().skip(1);
    let output = arguments
        .next()
        .ok_or("usage: create_iwork_chart_symbols <output-directory>")?;
    if arguments.next().is_some() {
        return Err("unexpected extra arguments".into());
    }
    let output = Path::new(&output);
    std::fs::create_dir_all(output)?;
    let symbols = [
        Some(ChartSeriesSymbol::sized(
            ChartSeriesSymbolShape::Square,
            ChartSeriesSymbolSize::new(18.0)?,
        )),
        Some(ChartSeriesSymbol::automatic(
            ChartSeriesSymbolShape::Diamond,
        )),
    ];
    let fills = [
        ChartSeriesSymbolFill::Custom(ShapeFill::Solid(RgbaColor::new(
            0.95,
            0.25,
            0.18,
            1.0,
            RgbColorSpace::Srgb,
        )?)),
        ChartSeriesSymbolFill::SeriesStroke,
    ];
    let outlines = [
        Some(ChartSeriesStroke::new(
            RgbaColor::black(),
            Width::new(2.5)?,
            ChartSeriesStrokePattern::RoundedDash,
        )),
        None,
    ];
    let connection_lines = [
        ChartSeriesConnectionLine::Straight,
        ChartSeriesConnectionLine::Curved,
    ];

    let mut numbers = NumbersDocumentBuilder::new()
        .sheet_name("Symbol CRUD")
        .build()?;
    let sheet_id = numbers.sheets()?[0].id();
    let chart = numbers.add_sheet_chart(
        sheet_id,
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 360.0, y: 100.0 },
        DrawableSize {
            width: 440.0,
            height: 300.0,
        },
    )?;
    numbers.set_sheet_chart_title(sheet_id, chart.drawable_object_id, "Typed data symbols")?;
    numbers.set_sheet_chart_series_symbols(sheet_id, chart.drawable_object_id, &symbols)?;
    numbers.set_sheet_chart_series_symbol_fills(sheet_id, chart.drawable_object_id, &fills)?;
    numbers.set_sheet_chart_series_symbol_outlines(
        sheet_id,
        chart.drawable_object_id,
        &outlines,
    )?;
    numbers.set_sheet_chart_series_connection_lines(
        sheet_id,
        chart.drawable_object_id,
        &connection_lines,
    )?;
    assert_eq!(
        numbers.sheet_chart_series_symbols(sheet_id, chart.drawable_object_id)?,
        symbols
    );
    assert_eq!(
        numbers.sheet_chart_series_symbol_fills(sheet_id, chart.drawable_object_id)?,
        fills
    );
    assert_eq!(
        numbers.sheet_chart_series_symbol_outlines(sheet_id, chart.drawable_object_id)?,
        outlines
    );
    assert_eq!(
        numbers.sheet_chart_series_connection_lines(sheet_id, chart.drawable_object_id)?,
        connection_lines
    );
    numbers.save(output.join("series-symbol.numbers"))?;

    let body = "Data symbol CRUD";
    let mut pages = PagesDocumentBuilder::new().body_text(body).build()?;
    let chart = pages.add_body_chart(
        body.encode_utf16().count(),
        Kind::Line2d,
        data()?,
        DrawablePoint { x: 72.0, y: 120.0 },
        DrawableSize {
            width: 420.0,
            height: 280.0,
        },
    )?;
    pages.set_body_chart_title(chart.drawable_object_id, "Typed data symbols")?;
    pages.set_body_chart_series_symbols(chart.drawable_object_id, &symbols)?;
    pages.set_body_chart_series_symbol_fills(chart.drawable_object_id, &fills)?;
    pages.set_body_chart_series_symbol_outlines(chart.drawable_object_id, &outlines)?;
    pages.set_body_chart_series_connection_lines(chart.drawable_object_id, &connection_lines)?;
    assert_eq!(
        pages.body_chart_series_symbols(chart.drawable_object_id)?,
        symbols
    );
    assert_eq!(
        pages.body_chart_series_symbol_fills(chart.drawable_object_id)?,
        fills
    );
    assert_eq!(
        pages.body_chart_series_symbol_outlines(chart.drawable_object_id)?,
        outlines
    );
    assert_eq!(
        pages.body_chart_series_connection_lines(chart.drawable_object_id)?,
        connection_lines
    );
    pages.save(output.join("series-symbol.pages"))?;

    let mut keynote = KeynoteDocumentBuilder::new()
        .title("Data Symbol CRUD")
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
    keynote.set_slide_chart_title(0, chart.drawable_object_id, "Typed data symbols")?;
    keynote.set_slide_chart_series_symbols(0, chart.drawable_object_id, &symbols)?;
    keynote.set_slide_chart_series_symbol_fills(0, chart.drawable_object_id, &fills)?;
    keynote.set_slide_chart_series_symbol_outlines(0, chart.drawable_object_id, &outlines)?;
    keynote.set_slide_chart_series_connection_lines(
        0,
        chart.drawable_object_id,
        &connection_lines,
    )?;
    assert_eq!(
        keynote.slide_chart_series_symbols(0, chart.drawable_object_id)?,
        symbols
    );
    assert_eq!(
        keynote.slide_chart_series_symbol_fills(0, chart.drawable_object_id)?,
        fills
    );
    assert_eq!(
        keynote.slide_chart_series_symbol_outlines(0, chart.drawable_object_id)?,
        outlines
    );
    assert_eq!(
        keynote.slide_chart_series_connection_lines(0, chart.drawable_object_id)?,
        connection_lines
    );
    keynote.save(output.join("series-symbol.key"))?;

    println!("created typed data-symbol fixtures in {}", output.display());
    Ok(())
}

fn data() -> Result<ChartData, Box<dyn std::error::Error>> {
    Ok(ChartData::new(
        vec!["North".to_owned(), "South".to_owned()],
        vec!["Q1".to_owned(), "Q2".to_owned(), "Q3".to_owned()],
        vec![
            vec![Some(12.0), Some(18.0), Some(24.0)],
            vec![Some(9.0), Some(21.0), Some(27.0)],
        ],
    )?)
}
