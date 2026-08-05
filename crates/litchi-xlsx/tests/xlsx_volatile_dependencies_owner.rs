use litchi_ooxml::xlsx::{
    VolatileDependencies, VolatileDependenciesConformance, Workbook,
    load_volatile_dependencies_from_package,
};

const MAIN_NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn fixture_value() -> VolatileDependencies {
    VolatileDependencies::parse(
        format!(
            r#"<volTypes xmlns="{MAIN_NS}"><volType type="realTimeData"><main first="server.id"><tp t="s"><v>ready</v><tr r="A1" s="0"/></tp></main></volType></volTypes>"#
        )
        .as_bytes(),
    )
    .expect("parse volatile dependencies")
}

#[test]
fn legacy_host_preserves_volatile_dependencies_through_writer_materialization() {
    let mut workbook = Workbook::create().expect("create workbook");
    let value = fixture_value();
    workbook
        .set_volatile_dependencies(&value, VolatileDependenciesConformance::Strict)
        .expect("set volatile dependencies");
    assert_eq!(
        workbook
            .volatile_dependencies()
            .expect("read volatile dependencies"),
        Some((value.clone(), VolatileDependenciesConformance::Strict))
    );
    assert_eq!(
        load_volatile_dependencies_from_package(workbook.opc_package())
            .expect("read forwarded volatile dependencies"),
        Some(value.clone())
    );

    workbook
        .worksheet_mut(0)
        .expect("first worksheet")
        .set_cell_value(1, 1, "materialized");
    let directory = tempfile::tempdir().expect("temporary directory");
    let path = directory
        .path()
        .join("materialized-volatile-dependencies.xlsx");
    workbook.save(&path).expect("save workbook");
    let reopened = Workbook::open(&path).expect("reopen workbook");
    assert_eq!(
        reopened
            .volatile_dependencies()
            .expect("read saved volatile dependencies"),
        Some((value.clone(), VolatileDependenciesConformance::Strict))
    );

    let mut reopened = reopened;
    assert!(
        reopened
            .remove_volatile_dependencies()
            .expect("remove volatile dependencies")
    );
    assert_eq!(
        reopened
            .volatile_dependencies()
            .expect("read removed volatile dependencies"),
        None
    );
    assert!(
        !reopened
            .remove_volatile_dependencies()
            .expect("idempotent volatile dependencies removal")
    );
}
