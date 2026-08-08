use std::sync::Arc;

use litchi_opc::OpcPackage;
use litchi_xlsx::calculation_properties::{
    Feature, Features, Limits, Properties, Transaction, apply_patch, parse_with_limits,
};
use litchi_xlsx::{Error, Package};

const NS: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const XCALCF: &str = "http://schemas.microsoft.com/office/spreadsheetml/2018/calcfeatures";
const FEATURES_URI: &str = "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}";

fn workbook(body: &str) -> Vec<u8> {
    format!(r#"<workbook xmlns="{NS}">{body}</workbook>"#).into_bytes()
}

fn features(names: &[&str]) -> Vec<u8> {
    let values = names
        .iter()
        .map(|name| format!(r#"<f:feature name="{name}"/>"#))
        .collect::<String>();
    workbook(&format!(
        r#"<extLst><ext uri="{FEATURES_URI}"><f:calcFeatures xmlns:f="{XCALCF}">{values}</f:calcFeatures></ext></extLst>"#,
    ))
}

fn nested_workbook(depth: usize) -> Vec<u8> {
    assert!(depth >= 1);
    let children = depth - 1;
    workbook(&format!(
        "{}{}",
        "<x>".repeat(children),
        "</x>".repeat(children)
    ))
}

fn package() -> OpcPackage {
    Package::create().expect("create package").into_plain_opc()
}

fn source(package: &OpcPackage) -> Arc<Vec<u8>> {
    package
        .main_document_part()
        .expect("workbook part")
        .blob_arc()
}

fn assert_limit_error(error: Error) {
    assert!(
        matches!(
            error,
            Error::Invalid(_) | Error::MarkupCompatibility(_) | Error::Xml(_)
        ),
        "unexpected typed error: {error}"
    );
}

#[test]
fn raw_and_mce_byte_limits_accept_exact_n_and_reject_n_plus_one() {
    let exact = workbook("");
    let over = workbook(" ");
    assert_eq!(over.len(), exact.len() + 1);

    let raw_limits = Limits::new().with_max_raw_bytes(exact.len()).unwrap();
    assert!(parse_with_limits(&exact, &raw_limits).is_ok());
    assert_limit_error(parse_with_limits(&over, &raw_limits).unwrap_err());

    let mce_limits = Limits::new().with_max_mce_bytes(exact.len()).unwrap();
    assert!(parse_with_limits(&exact, &mce_limits).is_ok());
    assert_limit_error(parse_with_limits(&over, &mce_limits).unwrap_err());
}

#[test]
fn depth_event_and_attribute_limits_accept_exact_n_and_reject_n_plus_one() {
    let depth_limits = Limits::new().with_max_depth(3).unwrap();
    assert!(parse_with_limits(&nested_workbook(3), &depth_limits).is_ok());
    assert_limit_error(parse_with_limits(&nested_workbook(4), &depth_limits).unwrap_err());

    let event_limits = Limits::new().with_max_events(3).unwrap();
    assert!(parse_with_limits(&workbook(""), &event_limits).is_ok());
    assert_limit_error(parse_with_limits(&workbook("<x/>"), &event_limits).unwrap_err());

    let attribute_limits = Limits::new().with_max_attributes(2).unwrap();
    assert!(parse_with_limits(&workbook(r#"<x a="1" b="2"/>"#), &attribute_limits).is_ok());
    assert_limit_error(
        parse_with_limits(&workbook(r#"<x a="1" b="2" c="3"/>"#), &attribute_limits).unwrap_err(),
    );
}

#[test]
fn feature_limits_accept_exact_n_and_reject_n_plus_one() {
    let count_limits = Limits::new().with_max_features(2).unwrap();
    assert!(parse_with_limits(&features(&["a", "b"]), &count_limits).is_ok());
    assert_limit_error(parse_with_limits(&features(&["a", "b", "c"]), &count_limits).unwrap_err());

    let name_limits = Limits::new().with_max_feature_name_bytes(3).unwrap();
    assert!(parse_with_limits(&features(&["abc"]), &name_limits).is_ok());
    assert_limit_error(parse_with_limits(&features(&["abcd"]), &name_limits).unwrap_err());

    let aggregate_limits = Limits::new().with_max_feature_names_bytes(4).unwrap();
    assert!(parse_with_limits(&features(&["ab", "cd"]), &aggregate_limits).is_ok());
    assert_limit_error(
        parse_with_limits(&features(&["ab", "cd", "e"]), &aggregate_limits).unwrap_err(),
    );
}

#[test]
fn opaque_byte_limit_accepts_exact_n_and_rejects_n_plus_one() {
    let limits = Limits::new().with_max_opaque_bytes(4).unwrap();
    assert!(parse_with_limits(&workbook("<extLst><x/></extLst>"), &limits).is_ok());
    assert_limit_error(
        parse_with_limits(&workbook("<extLst><xx/></extLst>"), &limits).unwrap_err(),
    );
}

fn authored_source(calculation_id: u32, limits: &Limits) -> Result<Vec<u8>, Error> {
    let mut package = package();
    let mut transaction = Transaction::with_limits(&mut package, limits)?;
    transaction.set_properties(Properties::new().with_calculation_id(Some(calculation_id)));
    transaction.commit()?;
    Ok(source(&package).as_ref().clone())
}

#[test]
fn output_limit_accepts_exact_n_and_failed_transaction_is_atomic_at_n_plus_one() {
    let exact_output = authored_source(1, &Limits::new()).unwrap();
    let over_output = authored_source(10, &Limits::new()).unwrap();
    assert_eq!(over_output.len(), exact_output.len() + 1);

    let limits = Limits::new()
        .with_max_output_bytes(exact_output.len())
        .unwrap();
    assert_eq!(authored_source(1, &limits).unwrap(), exact_output);

    let mut rejected = package();
    let before = source(&rejected);
    let mut transaction = Transaction::with_limits(&mut rejected, &limits).unwrap();
    transaction.set_properties(Properties::new().with_calculation_id(Some(10)));
    assert_limit_error(transaction.commit().unwrap_err());
    let after = source(&rejected);
    assert!(Arc::ptr_eq(&before, &after));
    assert_eq!(after.as_slice(), before.as_slice());
}

#[test]
fn failed_patch_is_atomic_and_retains_the_exact_source() {
    let mut authored = package();
    let mut transaction = Transaction::new(&mut authored).unwrap();
    transaction.set_properties(Properties::new().with_calculation_id(Some(7)));
    let patch = transaction.commit().unwrap().patch().clone();

    let mut stale = package();
    let mut transaction = Transaction::new(&mut stale).unwrap();
    transaction.set_properties(Properties::new().with_calculation_id(Some(8)));
    transaction.commit().unwrap();
    let before = source(&stale);

    assert!(matches!(
        apply_patch(&mut stale, &patch),
        Err(Error::PatchConflict { .. })
    ));
    let after = source(&stale);
    assert!(Arc::ptr_eq(&before, &after));
    assert_eq!(after.as_slice(), before.as_slice());
}

#[test]
fn every_limit_builder_rejects_zero_with_a_typed_error() {
    let limits = Limits::new();
    let results = [
        limits.with_max_raw_bytes(0),
        limits.with_max_mce_bytes(0),
        limits.with_max_output_bytes(0),
        limits.with_max_depth(0),
        limits.with_max_events(0),
        limits.with_max_attributes(0),
        limits.with_max_features(0),
        limits.with_max_feature_name_bytes(0),
        limits.with_max_feature_names_bytes(0),
        limits.with_max_opaque_bytes(0),
    ];

    for result in results {
        assert!(matches!(result, Err(Error::Invalid(_))));
    }
}

#[test]
fn staged_features_are_checked_by_writer_count_and_name_limits() {
    let scenarios = [
        (Limits::new().with_max_features(1).unwrap(), vec!["a", "b"]),
        (
            Limits::new().with_max_feature_name_bytes(1).unwrap(),
            vec!["ab"],
        ),
        (
            Limits::new().with_max_feature_names_bytes(1).unwrap(),
            vec!["a", "b"],
        ),
    ];

    for (limits, names) in scenarios {
        let mut package = package();
        let before = source(&package);
        let values = names
            .into_iter()
            .map(|name| Feature::new(name).unwrap())
            .collect::<Vec<_>>();
        let mut transaction = Transaction::with_limits(&mut package, &limits).unwrap();
        transaction.set_features(Features::try_from_vec(values).unwrap());
        assert_limit_error(transaction.commit().unwrap_err());
        assert!(Arc::ptr_eq(&before, &source(&package)));
    }
}
