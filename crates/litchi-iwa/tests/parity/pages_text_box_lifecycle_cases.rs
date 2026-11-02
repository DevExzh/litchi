//! Pages body text-box and shape lifecycle parity (wave W4-P, record 0413):
//! the host `remove_text_box` and `remove_body_shape` against
//! `litchi_pages::drawable::lifecycle::Edit::remove`.

use litchi_iwa::pages;
use litchi_pages::{DrawableSelector, Package};

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier};
use crate::leaf_bytes;
use crate::pages_drawable_cases::shape_fixture;
use crate::pages_text_box_cases::{host_text_box, text_box_fixture};

/// The body text and every drawable-owned text, through the host.
fn host_texts(bytes: &[u8]) -> TestResult<String> {
    let e = pages::Draft::from_bytes(bytes)?;
    let mut texts = e
        .drawable_text_storages()?
        .into_iter()
        .map(|item| item.storage.storage.into_text())
        .collect::<Vec<_>>();
    texts.sort();
    Ok(format!("{:?}\n{texts:?}", e.body_text()?))
}

/// The same through the leaf: body text, then every shape's and text box's
/// text.
fn leaf_texts(bytes: &[u8]) -> TestResult<String> {
    let package = Package::from_bytes(bytes)?;
    let mut texts = Vec::new();
    for make in [DrawableSelector::text_box, DrawableSelector::shape] {
        let mut index = 0;
        while let Ok(text) = package.body_drawable_text(make(index)) {
            texts.push(text);
            index += 1;
        }
    }
    texts.sort();
    let body = package.body_storage()?.text().to_owned();
    Ok(format!("{body:?}\n{texts:?}"))
}

fn leaf_remove(source: &[u8], selector: DrawableSelector) -> TestResult<LeafRun> {
    let package = Package::from_bytes(source)?;
    let mut edit = package.edit_body_media();
    edit.remove(selector)?;
    let commit = edit.commit()?;
    let restored = commit
        .package()
        .apply_body_media(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

fn remove_case(
    name: &'static str,
    input: &Fixture,
    host: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
    selector: DrawableSelector,
) -> TestResult {
    Case::new(name, Format::Pages, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .host(host)
        .leaf(|source| leaf_remove(source, selector))
        .read("host texts", host_texts)
        .read("leaf texts", leaf_texts)
        .run()?;
    Ok(())
}

#[test]
fn pages_text_box_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    let input = text_box_fixture()?;
    for index in [0, 1] {
        remove_case(
            "pages_text_box_remove_parity",
            &input,
            |source| {
                let mut e = pages::Draft::from_bytes(source)?;
                let d = host_text_box(&e, index)?;
                e.remove_text_box(d)?;
                Ok(e.to_bytes()?)
            },
            DrawableSelector::text_box(index),
        )
        .map_err(|e| format!("text box {index}: {e}"))?;
    }
    Ok(())
}

#[test]
fn pages_body_shape_remove_parity() -> TestResult {
    let _ = (Tier::B, Evidence::Queued);
    remove_case(
        "pages_body_shape_remove_parity",
        &shape_fixture()?,
        |source| {
            let mut e = pages::Draft::from_bytes(source)?;
            let d = e
                .body_shapes()?
                .first()
                .ok_or("no shape")?
                .drawable_object_id;
            e.remove_body_shape(d)?;
            Ok(e.to_bytes()?)
        },
        DrawableSelector::shape(0),
    )
}
