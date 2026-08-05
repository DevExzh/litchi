use litchi_xlsx::data_validation::{
    Collection, Conformance, Formula, ListSource, Source, Sqref, Validation, ValidationOperator,
    ValidationType, parse_data_validation_collections, validate_data_validation_collections,
    write_data_validation_collections,
};

fn values() -> Vec<Collection> {
    let mut validation = Validation::new(
        Source::Core,
        ValidationType::Whole,
        Sqref::parse("A1:B2").unwrap(),
    );
    validation.set_operator(ValidationOperator::Between);
    validation
        .set_formula1(Some(ListSource::Formula(Formula::new("1").unwrap())))
        .unwrap();
    validation
        .set_formula2(Some(Formula::new("10").unwrap()))
        .unwrap();
    vec![Collection::new(Source::Core, vec![validation]).unwrap()]
}

#[test]
fn data_validation_codec_round_trips_a_typed_collection() {
    let values = values();
    validate_data_validation_collections(&values).unwrap();
    let fragment = write_data_validation_collections(&values, Conformance::Transitional).unwrap();
    let xml = format!(
        r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">{fragment}</worksheet>"#
    );
    let parsed = parse_data_validation_collections(xml.as_bytes()).unwrap();
    assert_eq!(parsed.len(), 1);
    assert_eq!(parsed[0].validations().len(), 1);
    assert_eq!(
        parsed[0].validations()[0].sqref().ranges()[0].as_str(),
        "A1:B2"
    );
}

#[test]
fn invalid_sqref_is_rejected_before_serialization() {
    assert!(Sqref::parse("XFE1").is_err());
}
