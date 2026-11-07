//! Inert dependency hierarchy derivatives on a real five-sheet native workbook.
use litchi_xlsb::{
    Package,
    volatile_dependencies::{
        self as vol, CachedValue, CellReference, Dependencies, DependencyKind, ErrorCode,
        MainTopic, Topic, VolatileType,
    },
};
use std::collections::BTreeMap;
const FILE: &str = "excel-metadata.xlsb";
const OWNER: &str = "xl/volatileDependencies.bin";
fn native() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(FILE),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn changed(a: &[u8], b: &[u8]) -> Vec<String> {
    let a = members(a);
    let b = members(b);
    let names: std::collections::BTreeSet<_> = a.keys().chain(b.keys()).collect();
    names
        .into_iter()
        .filter(|n| a.get(*n) != b.get(*n))
        .cloned()
        .collect()
}

fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}

type Relationships = BTreeMap<String, (String, String, bool)>;
fn graph(p: &Package) -> BTreeMap<String, (String, Relationships)> {
    let opc = p.opc_package();
    let edges = |rels: &litchi_opc::Relationships| -> Relationships {
        rels.iter()
            .map(|r| {
                (
                    r.r_id().to_owned(),
                    (
                        r.reltype().to_owned(),
                        r.target_ref().to_owned(),
                        r.is_external(),
                    ),
                )
            })
            .collect()
    };
    let mut out = BTreeMap::new();
    out.insert("/".into(), (String::new(), edges(opc.rels())));
    for part in opc.iter_parts() {
        out.insert(
            part.partname().to_string(),
            (part.content_type().to_owned(), edges(part.rels())),
        );
    }
    out
}

fn sample() -> Dependencies {
    Dependencies::new(vec![
        VolatileType {
            kind: DependencyKind::Rtd,
            mains: vec![MainTopic {
                first: "Litchi.InertServer".into(),
                topics: vec![
                    Topic {
                        subtopics: vec!["Native source".into(), "価格".into()],
                        value: CachedValue::Number(30.),
                        references: vec![
                            CellReference::new(6, 0, 0).unwrap(),
                            CellReference::new(1, 1, 4).unwrap(),
                        ],
                    },
                    Topic {
                        subtopics: vec!["Second".into()],
                        value: CachedValue::Bool(false),
                        references: vec![CellReference::new(2, 1, 1).unwrap()],
                    },
                ],
            }],
        },
        VolatileType {
            kind: DependencyKind::Cube,
            mains: vec![MainTopic {
                first: "InertCubeConnection".into(),
                topics: vec![Topic {
                    subtopics: vec!["[Measures].[SavedValue]".into()],
                    value: CachedValue::String("unchanged".into()),
                    references: vec![CellReference::new(5, 2, 3).unwrap()],
                }],
            }],
        },
    ])
}
#[test]
fn native_workbook_volatile_owner_lifecycle_preserves_all_native_parts_and_inverses() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    assert!(!p.volatile_dependencies().unwrap().is_present());
    assert!(p.volatile_dependencies().unwrap().can_edit());
    let mut e = p.edit_volatile_dependencies().unwrap();
    assert!(!e.remove().unwrap());
    assert!(e.commit().unwrap().package().to_bytes().unwrap() == bytes);
    let value = sample();
    let mut e = p.edit_volatile_dependencies().unwrap();
    assert!(e.set_dependencies(value.clone()).unwrap());
    let created = e.commit().unwrap();
    let out = save(created.package());
    let after = out.to_bytes().unwrap();
    assert_eq!(
        changed(&bytes, &after),
        ["[Content_Types].xml", "xl/_rels/workbook.bin.rels", OWNER]
    );
    assert_eq!(
        out.volatile_dependencies().unwrap().dependencies(),
        Some(&value)
    );
    assert!(matches!(
        out.apply_volatile_dependencies(created.patch()),
        Err(vol::Error::PatchConflict { .. })
    ));
    assert!(
        members(
            &p.apply_volatile_dependencies(created.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&after)
    );
    assert!(
        members(
            &out.apply_volatile_dependencies(&created.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&bytes)
    );
    let mut e = out.edit_volatile_dependencies().unwrap();
    assert!(!e.set_dependencies(value.clone()).unwrap());
    assert!(e.commit().unwrap().package().to_bytes().unwrap() == after);
    let mut value = value;
    value.types[0].mains[0].topics[0].value = CachedValue::Number(31.);
    value.types[1].mains[0].topics[0].value = CachedValue::Error(ErrorCode::Na);
    let mut e = out.edit_volatile_dependencies().unwrap();
    e.set_dependencies(value.clone()).unwrap();
    let updated = e.commit().unwrap();
    let newer = save(updated.package());
    assert_eq!(changed(&after, &newer.to_bytes().unwrap()), [OWNER]);
    assert_eq!(
        newer.volatile_dependencies().unwrap().dependencies(),
        Some(&value)
    );
    let mut e = newer.edit_volatile_dependencies().unwrap();
    assert!(e.remove().unwrap());
    assert!(!e.remove().unwrap());
    let removed = e.commit().unwrap();
    let empty = save(removed.package());
    assert!(!empty.volatile_dependencies().unwrap().is_present());
    assert_eq!(
        changed(&bytes, &empty.to_bytes().unwrap()),
        ["[Content_Types].xml", "xl/_rels/workbook.bin.rels"]
    );
    assert_eq!(graph(&empty), graph(&p));
    assert!(
        members(
            &empty
                .apply_volatile_dependencies(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&newer.to_bytes().unwrap())
    );
    assert!(
        members(
            &newer
                .apply_volatile_dependencies(&updated.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
        ) == members(&after)
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_volatile_dependency_sheet_bounds_and_bad_drafts_leave_source_and_prior_staging_intact() {
    let bytes = native();
    let p = Package::from_bytes(&bytes).unwrap();
    let value = sample();
    let mut e = p.edit_volatile_dependencies().unwrap();
    e.set_dependencies(value.clone()).unwrap();
    let mut bad = value.clone();
    bad.types[0].mains[0].topics[0].value = CachedValue::Number(f64::NAN);
    assert!(matches!(
        e.set_dependencies(bad),
        Err(vol::Error::InvalidFormat(_))
    ));
    assert_eq!(e.dependencies().unwrap(), Some(&value));
    let mut bad = value.clone();
    bad.types[1].kind = bad.types[0].kind;
    assert!(e.set_dependencies(bad).is_err());
    assert_eq!(e.dependencies().unwrap(), Some(&value));
    let mut invalid = e.clone();
    let mut bad = value.clone();
    bad.types[0].mains[0].topics[0].references[0] = CellReference::new(0, 0, 5).unwrap();
    invalid.set_dependencies(bad).unwrap();
    assert!(matches!(
        invalid.commit(),
        Err(vol::Error::InvalidFormat(_))
    ));
    assert!(p.to_bytes().unwrap() == bytes);
    let mut limited = p
        .edit_volatile_dependencies_with_options(vol::EditOptions::default().limits(
            vol::SemanticLimits {
                max_types: 1,
                ..Default::default()
            },
        ))
        .unwrap();
    assert!(matches!(
        limited.set_dependencies(value.clone()),
        Err(vol::Error::Limit { .. })
    ));
    assert_eq!(limited.dependencies().unwrap(), None);
    let out = save(e.commit().unwrap().package());
    assert_eq!(
        out.volatile_dependencies().unwrap().dependencies(),
        Some(&value)
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
