//! Native exact-source checks for title/body placeholder visibility.

use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::WireView;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_keynote::{
    MAX_OBJECTS, MAX_SLIDES, MAX_TEXT_BYTES, MAX_TEXT_FRAGMENTS, MAX_TEXT_STORAGES, Package,
    ReadOptions, SemanticLimits, SlideSelector,
    slide::placeholder::{Error, Kind, State},
};

const SLIDE_MEMBER: &str = "Index/Slide-2652150.iwa";
const SLIDE_IDENTIFIER: u64 = 2_652_150;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const BODY_REFERENCE_FIELD: u32 = 6;
const OWNED_DRAWABLES_FIELD: u32 = 7;
const DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const SLIDE_NUMBER_REFERENCE_FIELD: u32 = 20;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn fixture() -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key")
}

fn rewrite_slide_payload(
    source: &[u8],
    rewrite: impl FnOnce(&[u8]) -> TestResult<Vec<u8>>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == SLIDE_MEMBER)
        .ok_or_else(|| io::Error::other("missing native slide member"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    let slide = archive
        .object_mut(SLIDE_IDENTIFIER)
        .ok_or_else(|| io::Error::other("missing native slide object"))?;
    let index = slide
        .messages
        .iter()
        .position(|message| message.type_ == SLIDE_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing native slide message"))?;
    let rewritten = rewrite(&slide.messages[index].data)?;
    slide.replace_message_preserving_header(
        index,
        RawMessage {
            type_: SLIDE_MESSAGE_TYPE,
            data: rewritten,
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(SLIDE_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn without_body_placeholder(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_slide_payload(source, |payload| {
        let view = WireView::parse(payload)?;
        let body_reference = view
            .fields()
            .find(|field| field.number() == BODY_REFERENCE_FIELD)
            .ok_or_else(|| io::Error::other("missing native body reference"))?
            .payload()
            .to_vec();
        let mut rewritten = Vec::with_capacity(payload.len());
        for field in view.fields() {
            let is_role = field.number() == BODY_REFERENCE_FIELD;
            let is_body_membership = matches!(
                field.number(),
                OWNED_DRAWABLES_FIELD | DRAWABLES_Z_ORDER_FIELD
            ) && field.payload() == body_reference;
            if !is_role && !is_body_membership {
                rewritten.extend_from_slice(field.raw());
            }
        }
        Ok(rewritten)
    })
}

fn with_mismatched_body_membership(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_slide_payload(source, |payload| {
        let view = WireView::parse(payload)?;
        let body_reference = view
            .fields()
            .find(|field| field.number() == BODY_REFERENCE_FIELD)
            .ok_or_else(|| io::Error::other("missing native body reference"))?
            .payload()
            .to_vec();
        let mut rewritten = Vec::with_capacity(payload.len());
        for field in view.fields() {
            let remove =
                field.number() == DRAWABLES_Z_ORDER_FIELD && field.payload() == body_reference;
            if !remove {
                rewritten.extend_from_slice(field.raw());
            }
        }
        Ok(rewritten)
    })
}

fn length_field(number: u32, payload: &[u8]) -> TestResult<Vec<u8>> {
    let mut output = Vec::with_capacity(payload.len().saturating_add(8));
    litchi_iwa_common::wire::append_length_delimited_field(&mut output, number, payload)?;
    Ok(output)
}

fn malformed_body_reference(source: &[u8], mode: u8) -> TestResult<Vec<u8>> {
    rewrite_slide_payload(source, |payload| {
        let view = WireView::parse(payload)?;
        let body = view
            .fields()
            .find(|field| field.number() == BODY_REFERENCE_FIELD)
            .ok_or_else(|| io::Error::other("missing native body reference"))?;
        let mut rewritten = payload.to_vec();
        match mode {
            0 => rewritten.extend_from_slice(&length_field(BODY_REFERENCE_FIELD, body.payload())?),
            1 => rewritten.extend_from_slice(&[0x35, 0, 0, 0, 0]),
            2 => {
                let mut external = body.payload().to_vec();
                external.extend_from_slice(&[0x18, 1]);
                rewritten = Vec::with_capacity(payload.len().saturating_add(4));
                for field in view.fields() {
                    if field.number() == BODY_REFERENCE_FIELD {
                        rewritten
                            .extend_from_slice(&length_field(BODY_REFERENCE_FIELD, &external)?);
                    } else {
                        rewritten.extend_from_slice(field.raw());
                    }
                }
            },
            _ => return Err(io::Error::other("unknown reference corruption mode").into()),
        }
        Ok(rewritten)
    })
}

fn malformed_slide_number_reference(source: &[u8], mode: u8) -> TestResult<Vec<u8>> {
    rewrite_slide_payload(source, |payload| {
        let view = WireView::parse(payload)?;
        let number = view
            .fields()
            .find(|field| field.number() == SLIDE_NUMBER_REFERENCE_FIELD)
            .ok_or_else(|| io::Error::other("missing native slide-number reference"))?;
        let mut rewritten = payload.to_vec();
        match mode {
            0 => rewritten.extend_from_slice(&length_field(
                SLIDE_NUMBER_REFERENCE_FIELD,
                number.payload(),
            )?),
            1 => rewritten.extend_from_slice(&[0xa0, 1, 0]),
            2 => {
                let mut external = number.payload().to_vec();
                external.extend_from_slice(&[0x18, 1]);
                rewritten.clear();
                for field in view.fields() {
                    if field.number() == SLIDE_NUMBER_REFERENCE_FIELD {
                        rewritten.extend_from_slice(&length_field(
                            SLIDE_NUMBER_REFERENCE_FIELD,
                            &external,
                        )?);
                    } else {
                        rewritten.extend_from_slice(field.raw());
                    }
                }
            },
            _ => return Err(io::Error::other("unknown reference corruption mode").into()),
        }
        Ok(rewritten)
    })
}

#[test]
fn native_visible_noop_hide_apply_inverse_and_locality() -> TestResult<()> {
    let package = Package::open(fixture())?;
    let source = bytes(&package)?;
    for kind in [Kind::Title, Kind::Body] {
        assert_eq!(
            package.slide_placeholder_visibility(SlideSelector::index(0), kind)?,
            Some(State::Visible)
        );
    }
    let noop = package
        .edit_slide_placeholder_visibility(0usize, Kind::Title)?
        .set(State::Visible)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(bytes(noop.package())?, source);

    let commit = package
        .edit_slide_placeholder_visibility(0usize, Kind::Title)?
        .hide()
        .commit()?;
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit
            .package()
            .slide_placeholder_visibility(0usize, Kind::Title)?,
        Some(State::Hidden)
    );
    let target = bytes(commit.package())?;
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    let before = Catalog::from_bytes(&source)?;
    let after = Catalog::from_bytes(&target)?;
    for (left, right) in before.iter().zip(after.iter()) {
        if left.data() == right.data() {
            assert_eq!(
                left.raw_record().local_record(),
                right.raw_record().local_record()
            );
        }
    }
    let applied = package.apply_slide_placeholder_visibility(commit.patch())?;
    assert_eq!(bytes(applied.package())?, target);
    assert!(matches!(
        commit
            .package()
            .apply_slide_placeholder_visibility(commit.patch()),
        Err(Error::PatchConflict)
    ));

    let shown = commit
        .package()
        .edit_slide_placeholder_visibility(0usize, Kind::Title)?
        .show()
        .commit()?;
    assert_eq!(
        shown
            .package()
            .slide_placeholder_visibility(0usize, Kind::Title)?,
        Some(State::Visible)
    );
    let hidden_again = shown
        .package()
        .apply_slide_placeholder_visibility(&shown.patch().inverse())?;
    assert_eq!(bytes(hidden_again.package())?, target);

    let restored = commit
        .package()
        .apply_slide_placeholder_visibility(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn missing_body_is_none_for_reads_and_typed_for_edits() -> TestResult<()> {
    let native = Package::open(fixture())?;
    let package = Package::from_bytes(&without_body_placeholder(&bytes(&native)?)?)?;
    assert_eq!(
        package.slide_placeholder_visibility(0usize, Kind::Body)?,
        None
    );
    assert!(matches!(
        package.edit_slide_placeholder_visibility(0usize, Kind::Body),
        Err(Error::PlaceholderNotFound { .. })
    ));
    Ok(())
}

#[test]
fn mismatched_membership_is_invalid_without_normalizing_source() -> TestResult<()> {
    let native = Package::open(fixture())?;
    let malformed = with_mismatched_body_membership(&bytes(&native)?)?;
    let package = Package::from_bytes(&malformed)?;
    assert!(matches!(
        package.slide_placeholder_visibility(0usize, Kind::Body),
        Err(Error::InvalidSource)
    ));
    assert_eq!(bytes(&package)?, malformed);
    Ok(())
}

#[test]
fn nested_body_reference_corruption_fails_closed_and_preserves_bytes() -> TestResult<()> {
    let native = Package::open(fixture())?;
    let source = bytes(&native)?;
    for mode in 0..3 {
        let malformed = malformed_body_reference(&source, mode)?;
        let package = Package::from_bytes(&malformed)?;
        assert!(matches!(
            package.edit_slide_placeholder_visibility(0usize, Kind::Body),
            Err(Error::InvalidSource)
        ));
        assert_eq!(bytes(&package)?, malformed);
    }
    Ok(())
}

#[test]
fn reference_limit_is_charged_on_placeholder_edit_admission() -> TestResult<()> {
    let source = std::fs::read(fixture())?;
    let limits = SemanticLimits::new(
        MAX_OBJECTS,
        MAX_SLIDES,
        1,
        MAX_TEXT_STORAGES,
        MAX_TEXT_FRAGMENTS,
        MAX_TEXT_BYTES,
    )?;
    let package =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), limits))?;
    let before = bytes(&package)?;
    assert!(matches!(
        package.edit_slide_placeholder_visibility(0usize, Kind::Title),
        Err(Error::LimitExceeded {
            kind: litchi_keynote::slide::placeholder::LimitKind::References,
            observed: 2,
            maximum: 1
        })
    ));
    assert_eq!(bytes(&package)?, before);
    Ok(())
}

#[test]
fn native_slide_number_hidden_noop_show_apply_and_inverse_are_exact() -> TestResult<()> {
    let package = Package::open(fixture())?;
    let source = bytes(&package)?;
    assert_eq!(
        package.slide_placeholder_visibility(0usize, Kind::SlideNumber)?,
        Some(State::Hidden)
    );
    let noop = package
        .edit_slide_placeholder_visibility(0usize, Kind::SlideNumber)?
        .set(State::Hidden)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(bytes(noop.package())?, source);
    let shown = package
        .edit_slide_placeholder_visibility(0usize, Kind::SlideNumber)?
        .show()
        .commit()?;
    assert!(shown.diagnostics().changed());
    assert_eq!(shown.diagnostics().touched_components(), 2);
    assert_eq!(shown.diagnostics().deleted_previews(), 3);
    assert_eq!(
        shown
            .package()
            .slide_placeholder_visibility(0usize, Kind::SlideNumber)?,
        Some(State::Visible)
    );
    let target = bytes(shown.package())?;
    let applied = package.apply_slide_placeholder_visibility(shown.patch())?;
    assert_eq!(bytes(applied.package())?, target);
    let restored = shown
        .package()
        .apply_slide_placeholder_visibility(&shown.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn slide_number_role_reference_corruption_is_atomic() -> TestResult<()> {
    let native = Package::open(fixture())?;
    let source = bytes(&native)?;
    for mode in 0..3 {
        let malformed = malformed_slide_number_reference(&source, mode)?;
        let package = Package::from_bytes(&malformed)?;
        assert!(matches!(
            package.edit_slide_placeholder_visibility(0usize, Kind::SlideNumber),
            Err(Error::InvalidSource)
        ));
        assert_eq!(bytes(&package)?, malformed);
    }
    Ok(())
}
