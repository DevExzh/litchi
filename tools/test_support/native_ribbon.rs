//! Explicit inert Ribbon derivatives of native Office packages.
use litchi_ooxml_common::ribbon::{Family, Set, Version};
use litchi_opc::{
    BlobPart, OpcPackage, PackURI, SourceLimits, constants::relationship_type as rt,
    phys_pkg::OwnedPhysPkgReader,
};
use std::{collections::BTreeMap, path::Path};

pub(crate) enum Action {
    Put(Version, Vec<u8>),
    Remove(Family),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) struct Slot {
    part: PackURI,
    id: String,
    version: Version,
    xml: Vec<u8>,
}

#[derive(Debug, PartialEq, Eq)]
pub(crate) struct Snapshot {
    legacy: Option<Slot>,
    modern: Option<Slot>,
    effective: Option<Version>,
}

pub(crate) fn snapshot(set: Set<'_>) -> Snapshot {
    let owned = |ui: litchi_ooxml_common::ribbon::Ui<'_>| Slot {
        part: ui.part().clone(),
        id: ui.id().into(),
        version: ui.version(),
        xml: ui.xml().into(),
    };
    Snapshot {
        legacy: set.legacy().map(owned),
        modern: set.modern().map(owned),
        effective: set.effective().map(|ui| ui.version()),
    }
}

pub(crate) trait Host {
    type Package;
    type Patch;
    const SUFFIX: &'static str;
    fn open(bytes: &[u8]) -> Self::Package;
    fn opc(p: &Self::Package) -> OpcPackage;
    fn from_opc(p: OpcPackage) -> Self::Package;
    fn bytes(p: &Self::Package) -> Vec<u8>;
    fn save(p: &Self::Package, path: &Path);
    fn reopen(path: &Path) -> Self::Package;
    fn read(p: &Self::Package) -> Snapshot;
    fn edit(p: &Self::Package, actions: &[Action]) -> (Self::Package, Self::Patch);
    fn apply(p: &Self::Package, patch: &Self::Patch) -> Self::Package;
    fn inverse(patch: &Self::Patch) -> Self::Patch;
    fn noop(patch: &Self::Patch) -> bool;
    fn assert_stale(p: &Self::Package, patch: &Self::Patch);
}

pub(crate) fn xml(version: Version, label: &str) -> Vec<u8> {
    format!(r#"<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="{}" onLoad="NeverInvokeOnLoad"><!-- retained native-parent derivative -->
<ribbon><tabs><tab id="nativeTab" label="{label}"><group id="nativeGroup" label="Saved">
<button id="nativeButton" label="日本語 🍎 &amp; &lt;saved&gt;" image="rIdNative" onAction="NeverInvokeAction" getEnabled="NeverInvokeEnabled"/>
</group></tab></tabs></ribbon></customUI>"#, version.namespace()).into_bytes()
}

pub(crate) fn native(root: &str, file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(root)
            .join("../../test-data/microsoft-office/excel")
            .join(file),
    )
    .unwrap()
}

fn members<H: Host>(p: &H::Package) -> BTreeMap<String, Vec<u8>> {
    let reader = OwnedPhysPkgReader::from_bytes(H::bytes(p), SourceLimits::default()).unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let bytes = reader.read_member(&name).unwrap();
            (name, bytes)
        })
        .collect()
}

type Graph = BTreeMap<String, (String, Vec<(String, String, String, bool)>)>;
fn graph<H: Host>(p: &H::Package) -> Graph {
    let p = H::opc(p);
    let rels = |rels: &litchi_opc::Relationships| {
        let mut result = rels
            .iter()
            .map(|r| {
                (
                    r.r_id().into(),
                    r.reltype().into(),
                    r.target_ref().into(),
                    r.is_external(),
                )
            })
            .collect::<Vec<_>>();
        result.sort();
        result
    };
    let mut result = BTreeMap::from([("/".into(), (String::new(), rels(p.rels())))]);
    for part in p.iter_parts() {
        result.insert(
            part.partname().to_string(),
            (part.content_type().into(), rels(part.rels())),
        );
    }
    result
}

fn equivalent<H: Host>(a: &H::Package, b: &H::Package) {
    assert_eq!(graph::<H>(a), graph::<H>(b));
    let a = members::<H>(a);
    let b = members::<H>(b);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert!(b.get(&name) == Some(&bytes), "changed member {name}");
        }
    }
}

fn retained<H: Host>(base: &H::Package, after: &H::Package) {
    let b = members::<H>(after);
    for (name, bytes) in members::<H>(base) {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert!(b.get(&name) == Some(&bytes), "native member changed {name}");
        }
    }
    let remove_ribbon = |entry: &mut (String, Vec<(String, String, String, bool)>)| {
        entry.1.retain(|r| {
            r.1 != Family::Legacy.relationship() && r.1 != Family::Modern.relationship()
        })
    };
    let mut a = graph::<H>(base);
    let mut b = graph::<H>(after);
    remove_ribbon(a.get_mut("/").unwrap());
    remove_ribbon(b.get_mut("/").unwrap());
    for (name, entry) in a {
        assert_eq!(b.get(&name), Some(&entry), "native graph changed {name}");
    }
}

fn reopened<H: Host>(p: &H::Package) -> H::Package {
    let file = tempfile::NamedTempFile::with_suffix(H::SUFFIX).unwrap();
    H::save(p, file.path());
    let result = H::reopen(file.path());
    assert!(members::<H>(&result) == members::<H>(p));
    assert_eq!(H::read(&result), H::read(p));
    result
}

pub(crate) fn published<H: Host>(
    before: &H::Package,
    result: (H::Package, H::Patch),
) -> H::Package {
    equivalent::<H>(&H::apply(&result.0, &H::inverse(&result.1)), before);
    let after = reopened::<H>(&result.0);
    equivalent::<H>(&H::apply(before, &result.1), &after);
    equivalent::<H>(&H::apply(&after, &H::inverse(&result.1)), before);
    after
}

pub(crate) fn lifecycle<H: Host>(root: &str, files: &[&str]) {
    for file in files {
        let bytes = native(root, file);
        let p = H::open(&bytes);
        assert_eq!(H::read(&p).effective, None);
        let legacy = xml(Version::V2007, "Legacy");
        let modern = xml(Version::V2010, "Modern");
        let both = published::<H>(
            &p,
            H::edit(
                &p,
                &[
                    Action::Put(Version::V2007, legacy.clone()),
                    Action::Put(Version::V2010, modern.clone()),
                ],
            ),
        );
        let original = H::read(&both);
        assert_eq!(original.effective, Some(Version::V2010));
        assert_eq!(original.legacy.as_ref().unwrap().xml, legacy);
        assert_eq!(original.modern.as_ref().unwrap().xml, modern);
        assert_ne!(
            original.legacy.as_ref().unwrap().id,
            original.modern.as_ref().unwrap().id
        );
        retained::<H>(&p, &both);
        let noop = H::edit(&both, &[Action::Put(Version::V2010, modern)]);
        assert!(H::noop(&noop.1));
        assert!(members::<H>(&noop.0) == members::<H>(&both));
        let ui2 = xml(Version::Ui2, "New modern version");
        let result = H::edit(&both, &[Action::Put(Version::Ui2, ui2.clone())]);
        H::assert_stale(&result.0, &result.1);
        let updated = published::<H>(&both, result);
        let current = H::read(&updated);
        assert_eq!(current.legacy, original.legacy);
        assert_eq!(current.effective, Some(Version::Ui2));
        let actual = current.modern.unwrap();
        let expected = original.modern.unwrap();
        assert_eq!((&actual.part, &actual.id), (&expected.part, &expected.id));
        assert_eq!(actual.xml, ui2);
        let a = members::<H>(&both);
        let b = members::<H>(&updated);
        assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
        assert_eq!(
            a.keys()
                .filter(|n| a.get(*n) != b.get(*n))
                .map(String::as_str)
                .collect::<Vec<_>>(),
            [actual.part.as_str().trim_start_matches('/')]
        );
        let legacy_only = published::<H>(
            &updated,
            H::edit(&updated, &[Action::Remove(Family::Modern)]),
        );
        assert_eq!(H::read(&legacy_only).effective, Some(Version::V2007));
        assert_eq!(H::read(&legacy_only).modern, None);
        let absent = H::edit(&legacy_only, &[Action::Remove(Family::Modern)]);
        assert!(H::noop(&absent.1));
        let cleared = published::<H>(
            &legacy_only,
            H::edit(&legacy_only, &[Action::Remove(Family::Legacy)]),
        );
        assert_eq!(H::read(&cleared).effective, None);
        equivalent::<H>(&cleared, &p);
        assert_eq!(H::bytes(&p), bytes);
        assert_eq!(native(root, file), bytes);
    }
}

pub(crate) fn image_ownership<H: Host>(root: &str, file: &str) {
    let bytes = native(root, file);
    let base = H::open(&bytes);
    let added = published::<H>(
        &base,
        H::edit(
            &base,
            &[Action::Put(
                Version::V2010,
                xml(Version::V2010, "With images"),
            )],
        ),
    );
    let ui = H::read(&added).modern.unwrap();
    let mut opc = H::opc(&added);
    let native_name = PackURI::new("/xl/media/image1.png").unwrap();
    assert_eq!(
        opc.get_part(&native_name).unwrap().content_type(),
        "image/png"
    );
    let image_bytes = opc.get_part(&native_name).unwrap().blob().to_vec();
    let owned_name = PackURI::new("/customUI/owned-native-copy.png").unwrap();
    opc.add_part(Box::new(BlobPart::new(
        owned_name.clone(),
        "image/png".into(),
        image_bytes.clone(),
    )));
    let part = opc.get_part_mut(&ui.part).unwrap();
    part.rels_mut().add_relationship(
        rt::IMAGE.into(),
        native_name.relative_ref(ui.part.base_uri()),
        "rIdNative".into(),
        false,
    );
    part.rels_mut().add_relationship(
        rt::IMAGE.into(),
        owned_name.relative_ref(ui.part.base_uri()),
        "rIdOwned".into(),
        false,
    );
    let imaged = reopened::<H>(&H::from_opc(opc));
    retained::<H>(&base, &imaged);
    let removed = published::<H>(&imaged, H::edit(&imaged, &[Action::Remove(Family::Modern)]));
    equivalent::<H>(&removed, &base);
    assert_eq!(
        H::opc(&removed).get_part(&native_name).unwrap().blob(),
        image_bytes
    );
    assert!(H::opc(&removed).get_part(&owned_name).is_err());

    let mut opc = H::opc(&imaged);
    let workbook = opc.main_document_part().unwrap().partname().clone();
    opc.get_part_mut(&workbook)
        .unwrap()
        .rels_mut()
        .add_relationship(
            "urn:litchi:retained-ribbon".into(),
            ui.part.relative_ref(workbook.base_uri()),
            "rIdKeepRibbon".into(),
            false,
        );
    let shared = reopened::<H>(&H::from_opc(opc));
    let replacement = xml(Version::V2010, "Forked 日本語");
    let forked = published::<H>(
        &shared,
        H::edit(&shared, &[Action::Put(Version::V2010, replacement.clone())]),
    );
    let fork_ui = H::read(&forked).modern.unwrap();
    assert_ne!(fork_ui.part, ui.part);
    assert_eq!(fork_ui.id, ui.id);
    assert_eq!(fork_ui.xml, replacement);
    retained::<H>(&shared, &forked);
    let fork_opc = H::opc(&forked);
    let fork_part = fork_opc.get_part(&fork_ui.part).unwrap();
    assert_eq!(fork_part.rels().len(), 2);
    for (id, name) in [("rIdNative", &native_name), ("rIdOwned", &owned_name)] {
        assert_eq!(
            fork_part.rels().get(id).unwrap().target_partname().unwrap(),
            *name
        );
    }
    let removed = published::<H>(&forked, H::edit(&forked, &[Action::Remove(Family::Modern)]));
    assert_eq!(H::read(&removed).effective, None);
    let mut expected_graph = graph::<H>(&shared);
    expected_graph
        .get_mut("/")
        .unwrap()
        .1
        .retain(|r| r.1 != Family::Modern.relationship());
    assert_eq!(graph::<H>(&removed), expected_graph);
    let a = members::<H>(&shared);
    let b = members::<H>(&removed);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert!(
                b.get(&name) == Some(&data),
                "shared dependency changed {name}"
            );
        }
    }
    assert_eq!(H::bytes(&base), bytes);
    assert_eq!(native(root, file), bytes);
}

pub(crate) fn rejected_edit<H: Host>(
    root: &str,
    file: &str,
    edit: impl FnOnce(&H::Package, Vec<u8>) -> (H::Package, H::Patch),
) {
    let bytes = native(root, file);
    let p = H::open(&bytes);
    let valid = xml(Version::V2010, "Valid staged data");
    let result = edit(&p, valid.clone());
    let q = published::<H>(&p, result);
    assert_eq!(H::read(&q).modern.unwrap().xml, valid);
    assert_eq!(H::read(&q).legacy, None);
    retained::<H>(&p, &q);
    assert_eq!(H::bytes(&p), bytes);
    assert_eq!(native(root, file), bytes);
}
