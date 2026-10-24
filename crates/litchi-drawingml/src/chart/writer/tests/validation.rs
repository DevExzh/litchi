use super::super::write;
use crate::chart::Chart;

#[test]
fn chart_style_accepts_the_ecma_defined_range() {
    for style in [1, 48] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        write(&mut Vec::new(), &chart, &crate::chart::Limits::DEFAULT).unwrap();
    }
}

#[test]
fn chart_style_rejects_values_outside_the_ecma_defined_range() {
    for style in [0, 49] {
        let mut chart = Chart::new();
        chart.style = Some(style);
        let error = write(&mut Vec::new(), &chart, &crate::chart::Limits::DEFAULT).unwrap_err();
        assert!(matches!(
            error,
            crate::Error::Io(error) if error.kind() == std::io::ErrorKind::InvalidInput
        ));
    }
}
