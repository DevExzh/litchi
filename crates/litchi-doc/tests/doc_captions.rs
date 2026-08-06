use litchi_doc::captions::{
    AutoEntry, AutoTable, Definition, Format, Heading, Info, LabelTable, Location, Numbering,
    Separator, Tables,
};

#[test]
fn facade_exposes_contextual_caption_types_without_repeated_prefixes() {
    let info = Info::new(
        Location::Above,
        Some(Numbering::new(Heading::Level1, Separator::Period)),
        false,
        Format::Arabic,
    );
    let labels =
        LabelTable::try_new(vec![Definition::try_new("Figure".into(), info).unwrap()]).unwrap();
    let auto = AutoTable::try_new(vec![
        AutoEntry::try_new("Word.Picture.8".into(), 0).unwrap(),
    ])
    .unwrap();
    let tables = Tables::try_new(Some(labels.clone()), Some(auto.clone())).unwrap();

    assert_eq!(tables.labels(), Some(&labels));
    assert_eq!(tables.auto(), Some(&auto));
    assert_eq!(
        labels.to_bytes().unwrap(),
        LabelTable::parse_bytes(&labels.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
    assert_eq!(
        auto.to_bytes().unwrap(),
        AutoTable::parse_bytes(&auto.to_bytes().unwrap())
            .unwrap()
            .to_bytes()
            .unwrap()
    );
}
