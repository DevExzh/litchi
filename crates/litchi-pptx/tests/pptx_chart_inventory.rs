use litchi_opc::constants::relationship_type::{CHART, STRICT_CHART};
use litchi_opc::{OpcError, OpcPackage, PackURI};
use litchi_pptx::chart::{self, Chart, Series, Type};
use litchi_pptx::{Error, Package};

#[test]
fn package_inventory_reports_generated_native_chart_with_identity() {
    let reopened = generated_package();
    let presentation = reopened.presentation().unwrap();
    let slide = presentation.slides().unwrap().remove(0);
    let charts = slide.charts().unwrap();
    assert_eq!(charts.len(), 1);

    let chart = &charts[0];
    assert_eq!(chart.part().partname().as_str(), "/ppt/charts/chart1.xml");
    let info = chart.chart_info().unwrap();
    assert_eq!(info.chart_type, Type::Bar);
    assert_eq!(info.title.as_deref(), Some("Quarterly sales"));
    assert!(info.has_legend);
    assert!(slide
        .part()
        .part()
        .rels()
        .iter()
        .any(|relationship| relationship.reltype() == CHART));

    let discovered = slide
        .charts()
        .unwrap()
        .into_iter()
        .map(|chart| chart.chart_info().unwrap())
        .collect::<Vec<_>>();
    assert_eq!(discovered, vec![info]);
}

#[test]
fn package_inventory_accepts_strict_chart_relationships() {
    let mut package = generated_package();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
    let (relationship_id, target) = {
        let slide = package.opc().unwrap().get_part(&slide_name).unwrap();
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

    package = edit_package(package, |opc| {
            let slide = opc.get_part_mut(&slide_name).unwrap();
            assert!(slide.rels_mut().remove(&relationship_id).is_some());
            slide.rels_mut().add_relationship(
                STRICT_CHART.to_string(),
                target,
                relationship_id,
                false,
            );
    });

    let charts = package.presentation().unwrap().slides().unwrap()[0]
        .charts()
        .unwrap();
    assert_eq!(charts.len(), 1);
    assert_eq!(charts[0].part().partname().as_str(), "/ppt/charts/chart1.xml");
}

#[test]
fn package_inventory_rejects_missing_chart_targets() {
    let mut package = generated_package();
    let chart_name = PackURI::new("/ppt/charts/chart1.xml").unwrap();
    package = edit_package(package, |opc| {
        assert!(opc.remove_part(&chart_name));
    });

    let error = match package
        .presentation()
        .unwrap()
        .slides()
        .unwrap()[0]
        .charts()
    {
        Ok(_) => panic!("a removed chart target must not be discovered"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::PartNotFound(message)) if message.contains("/ppt/charts/chart1.xml")
    ));
}

fn generated_package() -> Package {
    let mut package = Package::new().unwrap();
    let chart = Chart::new(Type::Bar, 914400, 914400, 4572000, 2743200)
        .with_title("Quarterly sales")
        .add_series(
            Series::new("2026")
                .with_categories(vec!["Q1".into(), "Q2".into()])
                .with_values(vec![100.0, 150.0]),
        );

    {
        package.presentation_mut().unwrap().add_slide().unwrap();
    }
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    chart::add(&mut opc, "/ppt/slides/slide1.xml", &chart).unwrap();
    Package::from_opc_package(opc).unwrap()
}

fn edit_package(mut package: Package, edit: impl FnOnce(&mut OpcPackage)) -> Package {
    let bytes = package.to_bytes().unwrap();
    let mut opc = OpcPackage::from_bytes(&bytes).unwrap();
    edit(&mut opc);
    Package::from_opc_package(opc).unwrap()
}
