#![allow(
    clippy::expect_used,
    reason = "tests deliberately panic on unexpected validation failures"
)]

use super::*;
use crate::{Cell, Column, Rect, Row};

#[test]
fn defaults_are_spreadsheetml_compatible() {
    let view = View::default();

    assert_eq!(view.window, Window::new(0));
    assert_eq!(view.mode, Mode::Normal);
    assert_eq!(view.color, Color::DEFAULT);
    assert_eq!(
        view.display,
        Display {
            window_protection: false,
            show_formulas: false,
            grid_lines: true,
            row_column_headers: true,
            zero_values: true,
            right_to_left: false,
            ruler: true,
            outline_symbols: true,
            default_grid_color: true,
            white_space: true,
        }
    );
    assert_eq!(view.zoom.current, Scale::DEFAULT);
    assert_eq!(view.zoom.normal, None);
    assert_eq!(view.zoom.page_layout, None);
    assert_eq!(view.zoom.page_break_preview, None);
    assert_eq!(view.origin, Cell::new(Row::FIRST, Column::FIRST));
    assert_eq!(view.pane, None);
    assert!(!view.tab_selected);
    assert_eq!(view.selections, vec![Selection::default()]);

    let selection = Selection::default();
    assert_eq!(selection.position(), Position::TopLeft);
    assert_eq!(
        selection.active_cell(),
        Cell::new(Row::FIRST, Column::FIRST)
    );
    assert_eq!(selection.active_range(), 0);
    assert_eq!(
        selection.ranges(),
        [Rect::single(Cell::new(Row::FIRST, Column::FIRST))]
    );
}

#[test]
fn pane_configuration_is_optional_and_keeps_its_own_origin() {
    let pane = Pane::default();

    assert_eq!(pane.position, Position::TopLeft);
    assert_eq!(pane.state, State::Split);
    assert_eq!(pane.top_left, Cell::new(Row::FIRST, Column::FIRST));
    assert_eq!(
        View {
            pane: Some(pane),
            ..View::default()
        }
        .pane,
        Some(pane)
    );
}

#[test]
fn every_view_enum_variant_is_representable() {
    assert_ne!(Mode::Normal, Mode::PageBreakPreview);
    assert_ne!(Mode::PageBreakPreview, Mode::PageLayout);
    assert_ne!(Position::BottomRight, Position::TopRight);
    assert_ne!(Position::TopRight, Position::BottomLeft);
    assert_ne!(Position::BottomLeft, Position::TopLeft);
    assert_ne!(State::Split, State::Frozen);
    assert_ne!(State::Frozen, State::FrozenSplit);
}

#[test]
fn checked_scalar_boundaries_are_preserved() {
    assert_eq!(Color::new(0).map(Color::get), Ok(0));
    assert_eq!(Color::new(64).map(Color::get), Ok(64));
    assert!(matches!(Color::new(65), Err(Error::Color { value: 65 })));

    assert_eq!(Scale::new(10).map(Scale::get), Ok(10));
    assert_eq!(Scale::new(400).map(Scale::get), Ok(400));
    assert!(matches!(Scale::new(9), Err(Error::Scale { value: 9 })));
    assert!(matches!(Scale::new(401), Err(Error::Scale { value: 401 })));

    assert_eq!(Split::new(0.0).map(Split::get), Ok(0.0));
    assert_eq!(Split::new(12.5).map(Split::get), Ok(12.5));
}

#[test]
fn nonfinite_or_negative_splits_are_rejected() {
    for value in [f64::NAN, -1.0, f64::INFINITY, f64::NEG_INFINITY] {
        assert!(matches!(Split::new(value), Err(Error::Split { .. })));
    }
}

#[test]
fn selections_require_ranges_and_an_in_bounds_active_range() {
    let cell = Cell::new(Row::FIRST, Column::FIRST);
    let range = Rect::single(cell);

    assert!(matches!(
        Selection::new(Position::TopLeft, cell, 0, Vec::new()),
        Err(Error::EmptySelection)
    ));
    assert!(matches!(
        Selection::new(Position::TopLeft, cell, 1, vec![range]),
        Err(Error::ActiveRange {
            active_range: 1,
            range_count: 1
        })
    ));
    assert_eq!(
        Selection::new(Position::BottomRight, cell, 0, vec![range])
            .expect("valid selection")
            .ranges(),
        [range]
    );
}

#[test]
fn last_grid_cell_can_be_the_active_selection() {
    let last = Cell::new(Row::LAST, Column::LAST);
    let selection = Selection::new(Position::BottomRight, last, 0, vec![Rect::single(last)])
        .expect("last grid cell selection");

    assert_eq!(selection.active_cell(), last);
    assert_eq!(selection.ranges(), [Rect::single(last)]);
    assert_eq!(last, Cell::new(Row::LAST, Column::LAST));
}
