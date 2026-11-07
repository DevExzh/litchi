//! Shared assertions for explicit add-in metadata derivatives of native Office parents.
use litchi_ooxml_common::web::{
    self as web, AddIn, Binding, Conformance, Pane, Panes, Property, Reference, Store,
};
use litchi_opc::{OpcPackage, SourceLimits, phys_pkg::OwnedPhysPkgReader};
use std::{collections::BTreeMap, path::Path};

pub(crate) const FIRST: &str = "native-addin-066-a";
pub(crate) const SECOND: &str = "native-addin-066-b";
pub(crate) const OWN_IMAGE: &str = "/media/native-addin-066.png";
const LINK: &str = "https://example.invalid/inert-snapshot.png?x=1&y=2";
const OPAQUE: &str = "<v:retained xmlns:v=\"urn:litchi:native-addin:066\" value=\"A &amp; B\"/>";

pub(crate) trait Host {
    type Package;
    type Patch;
    fn open(bytes: &[u8]) -> Self::Package;
    fn bytes(p: &Self::Package) -> Vec<u8>;
    fn save(p: &Self::Package, path: &Path);
    fn reopen(path: &Path) -> Self::Package;
    fn panes(p: &Self::Package) -> Option<Panes>;
    fn put(
        p: &Self::Package,
        panes: Panes,
        conformance: Conformance,
    ) -> (Self::Package, Self::Patch);
    fn remove(p: &Self::Package) -> (Self::Package, Self::Patch);
    fn apply(p: &Self::Package, patch: &Self::Patch) -> Self::Package;
    fn inverse(patch: &Self::Patch) -> Self::Patch;
    fn noop(patch: &Self::Patch) -> bool;
    fn assert_stale(p: &Self::Package, patch: &Self::Patch);
}

pub(crate) fn native(root: &str, file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(root)
            .join("../../test-data/microsoft-office")
            .join(file),
    )
    .unwrap()
}
pub(crate) fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let r = OwnedPhysPkgReader::from_bytes(bytes.to_vec(), SourceLimits::default()).unwrap();
    r.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = r.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn opc(bytes: &[u8]) -> OpcPackage {
    OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap()
}

pub(crate) fn panes(image: Option<&[u8]>) -> Panes {
    let reference = Reference::new("catalog-id-066", "1.2.3.4", Store::Omex)
        .unwrap()
        .location("en-US")
        .unwrap();
    let mut add_in = AddIn::new(FIRST, reference)
        .unwrap()
        .frozen(true)
        .bind(Binding::new("selection-066", "matrix", "native-app-ref-066").unwrap())
        .unwrap()
        .prop(Property::new("state", "日本語 & <saved>").unwrap())
        .unwrap();
    add_in
        .push_reference(
            Reference::file("fallback-066", "1.2.3.4", "file:///never-opened/addin.xml").unwrap(),
        )
        .unwrap();
    let xml = format!(
        "<we:extLst xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:ext uri=\"urn:litchi:opaque:066\">{OPAQUE}</a:ext></we:extLst>"
    );
    add_in
        .set_ext(web::ExtList::from_xml(xml.as_bytes()).unwrap())
        .unwrap();
    let mut functions = web::CustomFunctions::new();
    let mut ids = web::CustomFunctionList::new();
    ids.push_id("NATIVE.ADD")
        .unwrap()
        .push_id(" whitespace\t<&>🍎 ")
        .unwrap();
    functions
        .set_contains_custom_functions(Some(web::ContainsCustomFunctions::new(None)))
        .set_background_app_data(Some(
            web::BackgroundAppData::new(-7, "runtime\t<&>🍎").unwrap(),
        ))
        .set_custom_function_list(Some(ids));
    add_in.set_custom_functions(Some(functions)).unwrap();
    let mut a = Pane::new(add_in)
        .show(false)
        .dock(web::Dock::Left)
        .unwrap()
        .width(420.5)
        .unwrap()
        .linked(LINK)
        .unwrap();
    a.set_row(2).set_locked(true);
    let mut b = Pane::new(
        AddIn::new(
            SECOND,
            Reference::new("second-catalog", "2", Store::WopiCatalog).unwrap(),
        )
        .unwrap(),
    );
    if let Some(bytes) = image {
        a.set_image(OWN_IMAGE, "image/png", bytes.to_vec()).unwrap();
        b.set_image(OWN_IMAGE, "image/png", bytes.to_vec()).unwrap();
    }
    let mut panes = Panes::new();
    panes.push(a).unwrap().push(b).unwrap();
    panes
}

fn assert_model(actual: &Panes, expected: &Panes) {
    assert_eq!(actual.len(), expected.len());
    for (a, b) in actual.iter().zip(expected.iter()) {
        assert_eq!(
            (
                a.visible(),
                a.dock_state(),
                a.pane_width(),
                a.row(),
                a.locked()
            ),
            (
                b.visible(),
                b.dock_state(),
                b.pane_width(),
                b.row(),
                b.locked()
            )
        );
        assert_eq!(a.add_in(), b.add_in());
        assert_eq!(a.ext(), b.ext());
        assert_eq!(
            a.link().and_then(|l| l.external()),
            b.link().and_then(|l| l.external())
        );
        assert_eq!(
            a.image()
                .map(|i| (i.name().to_string(), i.content_type().to_owned())),
            b.image()
                .map(|i| (i.name().to_string(), i.content_type().to_owned()))
        );
        if let (Some(a), Some(b)) = (a.image(), b.image()) {
            assert!(a.bytes() == b.bytes());
        }
    }
    let p = actual.get(FIRST).unwrap();
    let a = p.add_in();
    assert_eq!(
        (
            p.visible(),
            p.dock_state(),
            p.pane_width(),
            p.row(),
            p.locked()
        ),
        (false, "left", 420.5, 2, true)
    );
    assert_eq!(a.reference().id(), "catalog-id-066");
    assert_eq!(a.reference().location_name(), Some("en-US"));
    assert_eq!(a.properties()[0].value(), "日本語 & <saved>");
    assert_eq!(a.bindings()[0].app_ref(), "native-app-ref-066");
    assert_eq!(p.link().unwrap().external(), Some(LINK));
    let functions = a.custom_functions().unwrap();
    assert_eq!(
        functions
            .contains_custom_functions()
            .unwrap()
            .explicit_value(),
        None
    );
    assert!(!functions.contains_custom_functions().unwrap().value());
    assert_eq!(functions.background_app_data().unwrap().state(), -7);
    assert_eq!(
        functions.background_app_data().unwrap().runtime_id(),
        "runtime\t<&>🍎"
    );
    assert_eq!(
        functions.custom_function_list().unwrap().ids(),
        ["NATIVE.ADD", " whitespace\t<&>🍎 "]
    );
    assert!(a.ext().unwrap().xml().contains(OPAQUE));
}

fn rels(r: &litchi_opc::Relationships) -> Vec<(String, String, String, bool)> {
    let mut v = r
        .iter()
        .map(|r| {
            (
                r.r_id().to_owned(),
                r.reltype().to_owned(),
                r.target_ref().to_owned(),
                r.is_external(),
            )
        })
        .collect::<Vec<_>>();
    v.sort();
    v
}

fn preserved(original: &[u8], output: &[u8]) {
    let a = members(original);
    let b = members(output);
    for (name, data) in a {
        if name != "[Content_Types].xml" && name != "_rels/.rels" {
            assert!(b.get(&name) == Some(&data), "native member changed: {name}");
        }
    }
    let a = opc(original);
    let b = opc(output);
    let roots = |p: &OpcPackage| {
        rels(p.rels())
            .into_iter()
            .filter(|r| r.1 != web::raw::TASK_PANES_RELATIONSHIP)
            .collect::<Vec<_>>()
    };
    assert_eq!(roots(&a), roots(&b));
    for part in a.iter_parts() {
        let other = b.get_part(part.partname()).unwrap();
        assert_eq!(part.content_type(), other.content_type());
        assert_eq!(rels(part.rels()), rels(other.rels()));
    }
}

fn equivalent(a: &[u8], b: &[u8]) {
    let left = members(a);
    let right = members(b);
    assert_eq!(
        left.keys().collect::<Vec<_>>(),
        right.keys().collect::<Vec<_>>()
    );
    preserved(a, b);
    // DOCX's generic resource replay can normalize only these owned graph files.
    // Require the complete graph, including task-pane root bindings, to agree.
    assert_eq!(rels(opc(a).rels()), rels(opc(b).rels()));
}

fn publication<H: Host>(source: &H::Package, changed: &H::Package, patch: &H::Patch) -> H::Package {
    let file = tempfile::NamedTempFile::new().unwrap();
    H::save(changed, file.path());
    let reopened = H::reopen(file.path());
    assert!(H::panes(&reopened) == H::panes(changed));
    let target = members(&H::bytes(changed));
    assert!(members(&H::bytes(&reopened)) == target);
    equivalent(&H::bytes(changed), &H::bytes(&H::apply(source, patch)));
    let inverse = H::apply(&reopened, &H::inverse(patch));
    equivalent(&H::bytes(source), &H::bytes(&inverse));
    reopened
}

pub(crate) fn lifecycle<H: Host>(root: &str, files: &[(&str, Conformance)]) {
    for (file, conformance) in files {
        let original = native(root, file);
        let source = H::open(&original);
        let source_bytes = H::bytes(&source);
        assert!(H::panes(&source).is_none(), "{file}");
        let original_members = members(&original);
        let image = original_members
            .iter()
            .find(|(n, _)| n.contains("/media/") && n.ends_with(".png"));
        let expected = panes(image.map(|(_, b)| b.as_slice()));
        let (changed, patch) = H::put(&source, expected.clone(), *conformance);
        assert!(!H::noop(&patch));
        let added = publication::<H>(&source, &changed, &patch);
        preserved(&original, &H::bytes(&added));
        assert_model(&H::panes(&added).unwrap(), &expected);
        let (same, noop) = H::put(&added, H::panes(&added).unwrap(), *conformance);
        assert!(H::noop(&noop));
        assert!(H::bytes(&same) == H::bytes(&added));
        H::assert_stale(&added, &patch);

        let mut modified = H::panes(&added).unwrap();
        modified
            .edit(FIRST, |pane| {
                pane.set_visible(true).set_width(512.0)?;
                let add_in = pane.add_in_mut();
                add_in.upsert_property(Property::new("state", "updated & reviewed")?)?;
                let mut custom = add_in.custom_functions().unwrap().clone();
                custom.set_background_app_data(Some(web::BackgroundAppData::new(
                    i32::MAX,
                    "updated runtime",
                )?));
                add_in.set_custom_functions(Some(custom))?;
                Ok(())
            })
            .unwrap();
        let (changed, patch) = H::put(&added, modified.clone(), *conformance);
        let updated = publication::<H>(&added, &changed, &patch);
        assert!(H::panes(&updated).unwrap() == modified);
        preserved(&original, &H::bytes(&updated));
        let old_opc = opc(&H::bytes(&added));
        let new_opc = opc(&H::bytes(&updated));
        assert_eq!(rels(old_opc.rels()), rels(new_opc.rels()));
        // Scalar metadata changes retain the complete part and relationship identities.
        assert_eq!(
            old_opc
                .iter_parts()
                .map(|p| p.partname().to_string())
                .collect::<std::collections::BTreeSet<_>>(),
            new_opc
                .iter_parts()
                .map(|p| p.partname().to_string())
                .collect()
        );
        for part in old_opc.iter_parts() {
            assert_eq!(
                rels(part.rels()),
                rels(new_opc.get_part(part.partname()).unwrap().rels())
            );
        }
        assert!(
            H::panes(&updated)
                .unwrap()
                .get(FIRST)
                .unwrap()
                .add_in()
                .ext()
                .unwrap()
                .xml()
                .contains(OPAQUE)
        );
        modified.remove(FIRST).unwrap();
        let (changed, patch) = H::put(&updated, modified.clone(), *conformance);
        let remaining = publication::<H>(&updated, &changed, &patch);
        assert!(H::panes(&remaining).unwrap() == modified);
        if let Some((_, payload)) = image {
            assert_eq!(
                members(&H::bytes(&remaining))[OWN_IMAGE.trim_start_matches('/')],
                *payload
            );
        }
        let (changed, patch) = H::remove(&remaining);
        let cleared = publication::<H>(&remaining, &changed, &patch);
        assert!(H::panes(&cleared).is_none());
        preserved(&original, &H::bytes(&cleared));
        assert_eq!(
            members(&H::bytes(&cleared)).keys().collect::<Vec<_>>(),
            original_members.keys().collect::<Vec<_>>()
        );
        let (_, patch) = H::remove(&cleared);
        assert!(H::noop(&patch));
        assert!(H::bytes(&source) == source_bytes);
        assert!(native(root, file) == original);
    }
}

pub(crate) fn collision<H: Host>(
    root: &str,
    file: &str,
    stage: impl FnOnce(&H::Package, Panes, Panes, &str) -> (H::Package, H::Patch),
) {
    let original = native(root, file);
    let source = H::open(&original);
    let source_bytes = H::bytes(&source);
    let all = members(&original);
    let (name, image) = all
        .iter()
        .find(|(n, _)| n.contains("/media/") && n.ends_with(".png"))
        .unwrap();
    let target = format!("/{name}");
    let valid = panes(Some(image));
    let mut invalid = valid.clone();
    invalid
        .edit(FIRST, |pane| {
            pane.set_image(&target, "image/png", image.clone())?;
            Ok(())
        })
        .unwrap();
    let reason =
        format!("authored web extension part '{target}' already exists outside the replaced graph");
    let (changed, patch) = stage(&source, valid.clone(), invalid, &reason);
    let retained = publication::<H>(&source, &changed, &patch);
    assert_model(&H::panes(&retained).unwrap(), &valid);
    preserved(&original, &H::bytes(&retained));
    assert!(H::bytes(&source) == source_bytes);
    assert!(native(root, file) == original);
}
