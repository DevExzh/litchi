//! Native save/reopen regression for shared grid authoring and complete replacement.

use litchi_keynote::{ChartData, ChartSelector, Package, SlideSelector};

#[test]
fn authored_and_replaced_grids_survive_native_save() -> Result<(), Box<dyn std::error::Error>> {
    for (stem, replaced) in [
        ("chart-grid-authored-native", false),
        ("chart-grid-replaced-native", true),
    ] {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/keynote")
            .join(format!("{stem}.key"));
        let source = std::fs::read(path)?;
        let package = Package::from_bytes(&source)?;
        let actual = package.slide_chart_data(SlideSelector::index(0), ChartSelector::index(0))?;
        let (columns, values) = if replaced {
            (
                vec!["April", "May", "June", "July"],
                vec![
                    vec![Some(27.5), None, Some(53.0), Some(96.0)],
                    vec![Some(55.0), Some(12.75), Some(70.0), Some(58.0)],
                ],
            )
        } else {
            (
                vec!["Q1", "Q2", "Q3"],
                vec![
                    vec![Some(12.0), Some(18.0), Some(24.0)],
                    vec![Some(9.0), Some(21.0), Some(27.0)],
                ],
            )
        };
        let expected = ChartData::new(
            vec!["North".into(), "South".into()],
            columns.into_iter().map(str::to_owned).collect(),
            values,
        )?;
        assert!(actual.bitwise_eq(&expected), "{stem}");
        assert_eq!(
            package
                .slide_chart_title(SlideSelector::index(0), ChartSelector::index(0))?
                .as_deref(),
            Some("Grid authoring proof"),
            "{stem}"
        );
        let mut preserved = Vec::new();
        package.write_to(&mut preserved)?;
        assert_eq!(preserved, source, "{stem}");
    }
    Ok(())
}
