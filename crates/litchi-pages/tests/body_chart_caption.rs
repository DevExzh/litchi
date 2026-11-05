//! Selector, reversibility and refusal coverage for body-chart captions
//! (record 0427).

use litchi_pages::chart::caption::{Error, Refusal};
use litchi_pages::{BodyChartSelector, Package};

const NATIVE: &[u8] = include_bytes!("../../../test-data/iwork/pages/chart-data-native.pages");

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn body_chart_caption_set_replace_clear_and_inverse() -> TestResult {
    let chart = BodyChartSelector::index(0);
    let package = Package::from_bytes(NATIVE)?;
    assert_eq!(package.body_chart_caption(chart)?, None);

    let mut edit = package.edit_body_chart_caption(chart)?;
    assert_eq!(edit.before(), None);
    edit.set("Revenue — 2026")?;
    let created = edit.commit()?;
    assert!(created.diagnostics().changed());
    assert_eq!(created.diagnostics().created_objects(), 4);
    assert_eq!(
        created.package().body_chart_caption(chart)?.as_deref(),
        Some("Revenue — 2026")
    );
    let restored = created
        .package()
        .apply_body_chart_caption(&created.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, NATIVE);

    let mut edit = created.package().edit_body_chart_caption(chart)?;
    edit.set("Updated")?;
    let replaced = edit.commit()?;
    assert_eq!(replaced.diagnostics().created_objects(), 0);
    assert_eq!(
        replaced.package().body_chart_caption(chart)?.as_deref(),
        Some("Updated")
    );

    let mut edit = replaced.package().edit_body_chart_caption(chart)?;
    edit.clear()?;
    let cleared = edit.commit()?;
    assert_eq!(cleared.diagnostics().created_objects(), 1);
    assert_eq!(cleared.package().body_chart_caption(chart)?, None);
    Ok(())
}

#[test]
fn body_chart_caption_refusals_are_typed() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    match package.body_chart_caption(BodyChartSelector::index(9)) {
        Err(Error::Refused(Refusal::NotFound { kind: "chart", .. })) => {},
        other => return Err(format!("expected NotFound, got {other:?}").into()),
    }
    let mut edit = package.edit_body_chart_caption(0)?;
    match edit.set("a\u{fffc}b") {
        Err(Error::Refused(Refusal::ReservedMarker)) => {},
        other => return Err(format!("expected ReservedMarker, got {other:?}").into()),
    }
    edit.set("ok")?;
    match edit.clear() {
        Err(Error::Refused(Refusal::OperationAlreadyStaged)) => {},
        other => return Err(format!("expected OperationAlreadyStaged, got {other:?}").into()),
    }
    let unstaged = package.edit_body_chart_caption(0)?.commit()?;
    assert!(unstaged.patch().is_noop());
    assert_eq!(exact_bytes(unstaged.package())?, NATIVE);
    Ok(())
}

/// No Pages file we hold has a caption Pages authored, so a new caption keeps
/// the theme's first preset, the title's style, and this pins that choice: the
/// template does list a caption style, and using it would be a guess until a
/// native file shows it (record 0435).
#[test]
fn new_caption_keeps_the_theme_first_preset_without_native_evidence() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    let mut edit = package.edit_body_chart_caption(BodyChartSelector::index(0))?;
    edit.set("Pinned caption")?;
    let written = exact_bytes(edit.commit()?.package())?;
    let catalog = litchi_iwa_archive::package::Catalog::from_bytes(&written, Default::default())?;
    let mut objects = Vec::new();
    for entry in catalog.iter() {
        if entry.name().ends_with(".iwa") {
            let stream = litchi_iwa_archive::iwa::SnappyStream::decompress(
                entry.data(),
                Default::default(),
            )?;
            objects.extend(
                litchi_iwa_archive::iwa::Archive::parse(stream.as_bytes(), Default::default())?
                    .objects,
            );
        }
    }
    let identifier = objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == 2_001)
        .find_map(|message| {
            message
                .data
                .windows(b"Pinned caption".len())
                .any(|window| window == b"Pinned caption")
                .then(|| paragraph_reference(&message.data))
        })
        .ok_or("no storage holds the caption")??;
    let style = objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(identifier))
        .ok_or("no paragraph style")?;
    let stored = &style.messages[0].data;
    let needle = b"captions-0-paragraphstyle-Caption Title";
    assert!(stored.windows(needle.len()).any(|window| window == needle));
    Ok(())
}

/// The object at `table_para_style.entries[0].object` of a text storage.
fn paragraph_reference(storage: &[u8]) -> TestResult<u64> {
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::wire::WireView;
    let mut payload = storage;
    for number in [5, 1, 2, 1] {
        let view = WireView::parse(payload, WireLimits::DEFAULT)?;
        payload = view
            .fields()
            .find(|field| field.number() == number)
            .map(|field| field.payload())
            .ok_or("a missing field")?;
    }
    Ok(litchi_iwa_common::decode_varint_from_bytes(payload)?.0)
}
