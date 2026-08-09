#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic by design"
)]

use super::super::super::{Kind, RowCol, cache};
use super::super::{Cache, Value};
use super::validation::{cache_dimensions, dimensions_cover};

#[test]
fn excel_cache_dimensions_are_derived_from_the_used_range() {
    let values = [
        Cache::excel(
            cache::Index::Values,
            4,
            2,
            cache::Xf::new(1),
            Value::Number(1.0),
        ),
        Cache::excel(
            cache::Index::Values,
            7,
            5,
            cache::Xf::new(2),
            Value::Number(2.0),
        ),
    ];
    let derived = cache_dimensions(&values, Kind::Excel).expect("Excel dimensions");
    let cache::Dims::Excel(excel_dims) = derived else {
        panic!("expected Excel dimensions");
    };
    assert_eq!(excel_dims.first_row(), 4);
    assert_eq!(excel_dims.row_after(), 8);
    assert_eq!(excel_dims.first_col(), 2);
    assert_eq!(excel_dims.col_after(), 6);
    assert!(dimensions_cover(
        cache::Dims::Excel(cache::ExcelDims::new(0, 10, 0, 8).expect("covering range")),
        cache::Dims::Excel(excel_dims),
    ));
    assert!(dimensions_cover(
        cache::Dims::Excel(cache::ExcelDims::new(0, 10, 0, 8).expect("declared source range")),
        cache::Dims::Excel(cache::ExcelDims::default()),
    ));
}

#[test]
fn graph_cache_dimensions_deduplicate_coordinates() {
    let values = [
        Cache::graph(
            RowCol::new(2).expect("row"),
            RowCol::new(3).expect("column"),
            cache::Ifmt::new(1),
            Value::Blank,
        ),
        Cache::graph(
            RowCol::new(2).expect("row"),
            RowCol::new(3).expect("column"),
            cache::Ifmt::new(1),
            Value::Blank,
        ),
        Cache::graph(
            RowCol::new(4).expect("row"),
            RowCol::new(1).expect("column"),
            cache::Ifmt::new(1),
            Value::Blank,
        ),
    ];
    let cache::Dims::Graph(derived) =
        cache_dimensions(&values, Kind::Graph).expect("Graph dimensions")
    else {
        panic!("expected Graph dimensions");
    };
    assert_eq!(derived.longest_row().get(), 1);
    assert_eq!(derived.rows(), 2);
}

#[test]
fn cache_dimensions_reject_cross_producer_values() {
    let graph = [Cache::graph(
        RowCol::new(1).expect("row"),
        RowCol::new(1).expect("column"),
        cache::Ifmt::new(0),
        Value::Blank,
    )];
    assert!(cache_dimensions(&graph, Kind::Excel).is_err());
    assert!(!dimensions_cover(
        cache::Dims::Excel(cache::ExcelDims::default()),
        cache::Dims::Graph(cache::GraphDims::default()),
    ));
}
