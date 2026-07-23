use litchi_ooxml::pptx::{ChartData, ChartSeries, ChartType, Package};
use litchi_ooxml::{OoxmlError, PackURI};
use litchi_opc::constants::relationship_type::{CHART, STRICT_CHART};
use tempfile::NamedTempFile;

#[test]
fn package_inventory_reports_generated_native_chart_with_identity() {
    let reopened = generated_package();
    let charts = reopened.charts().unwrap();
    assert_eq!(charts.len(), 1);

    let chart = &charts[0];
    assert_eq!(chart.slide_index(), 0);
    assert!(chart.relationship_id().starts_with("rId"));
    assert_eq!(chart.part_name().as_str(), "/ppt/charts/chart1.xml");
    assert_eq!(chart.info().chart_type, ChartType::Bar);
    assert_eq!(chart.info().title.as_deref(), Some("Quarterly sales"));
    assert!(chart.info().has_legend);

    let legacy = reopened.presentation().unwrap().get_charts().unwrap();
    assert_eq!(legacy, vec![(0, chart.info().clone())]);
}

#[test]
fn package_inventory_accepts_strict_chart_relationships() {
    let mut package = generated_package();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let (relationship_id, target) = {
        let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
        let relationship = slide
            .rels()
            .iter()
            .find(|relationship| relationship.reltype() == CHART)
            .unwrap();
        (
            relationship.r_id().to_string(),
            relationship.target_ref().to_string(),
        )
    };

    let slide = package.opc_package_mut().get_part_mut(&slide_name).unwrap();
    assert!(slide.rels_mut().remove(&relationship_id).is_some());
    slide
        .rels_mut()
        .add_relationship(STRICT_CHART.to_string(), target, relationship_id, false);

    let charts = package.charts().unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].part_name().as_str(), "/ppt/charts/chart1.xml");
}

#[test]
fn package_inventory_rejects_missing_chart_targets() {
    let mut package = generated_package();
    let chart_name = PackURI::new("/ppt/charts/chart1.xml").unwrap();
    assert!(package.opc_package_mut().remove_part(&chart_name));

    let error = package.charts().unwrap_err();
    assert!(matches!(
        error,
        OoxmlError::PartNotFound(message) if message.contains("/ppt/charts/chart1.xml")
    ));
}

fn generated_package() -> Package {
    let output = NamedTempFile::with_suffix(".pptx").unwrap();
    let mut package = Package::new().unwrap();
    let chart = ChartData::new(ChartType::Bar, 914400, 914400, 4572000, 2743200)
        .with_title("Quarterly sales")
        .add_series(
            ChartSeries::new("2026")
                .with_categories(vec!["Q1".into(), "Q2".into()])
                .with_values(vec![100.0, 150.0]),
        );

    {
        let presentation = package.presentation_mut().unwrap();
        let chart_index = presentation.add_chart_parts(&chart).unwrap();
        let slide = presentation.add_slide().unwrap();
        slide.add_chart_shape(chart_index, chart.x, chart.y, chart.width, chart.height);
    }
    package.save(output.path()).unwrap();

    Package::open(output.path()).unwrap()
}
