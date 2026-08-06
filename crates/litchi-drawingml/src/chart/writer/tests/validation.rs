use super::super::write;
use crate::chart::Chart;

#[test]
fn chart_style_accepts_the_ecma_defined_range() {
    for style in [1, 48] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        write(&mut Vec::new(), &chart).unwrap();
    }
}

#[test]
fn chart_style_rejects_values_outside_the_ecma_defined_range() {
    for style in [0, 49] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        let error = write(&mut Vec::new(), &chart).unwrap_err();
        assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    }
}
