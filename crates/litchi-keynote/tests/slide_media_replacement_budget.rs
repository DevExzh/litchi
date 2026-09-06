//! Aggregate-budget regression coverage for native Keynote slide media.
//!
//! Physical ingress and the lazy semantic document projection each admit the
//! native fixture under the edge reference profile below.  A media edit has a
//! larger operation-local closure: it scans the package catalog, metadata
//! ownership, wire records, and candidate reassembly.  The focused owner must
//! reject that transaction before publication when the same reference ceiling
//! is exhausted, while the immutable source remains byte-identical.

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_keynote::{
    MediaPart, MovieSelector, Package, ReadOptions, SemanticLimits, SlideMediaDataError,
    SlideMediaDataLimitKind, SlideSelector,
};

const NATIVE_SOURCE: &[u8] =
    include_bytes!("../../../test-data/iwork/keynote/media-replacement-native.key");
const AUDIO_MEMBER: &str = "Data/keynote-coral-9075.wav";
const MAX_REFERENCE_SEARCH_STEPS: usize = usize::BITS as usize + 1;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

fn member_bytes(name: &str) -> TestResult<Vec<u8>> {
    Catalog::from_bytes(NATIVE_SOURCE)?
        .iter()
        .find(|entry| entry.name() == name)
        .map(|entry| entry.data().to_vec())
        .ok_or_else(|| io::Error::other(format!("missing native member {name}")))
        .map_err(Into::into)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn package_with_reference_limit(limit: usize) -> TestResult<Package> {
    let semantic = SemanticLimits::new(
        SemanticLimits::MAX_OBJECTS,
        SemanticLimits::MAX_SLIDES,
        limit,
        SemanticLimits::MAX_TEXT_STORAGES,
        SemanticLimits::MAX_TEXT_FRAGMENTS,
        SemanticLimits::MAX_TEXT_BYTES,
    )?;
    Ok(Package::from_bytes_with_options(
        NATIVE_SOURCE,
        ReadOptions::new(Limits::default(), semantic),
    )?)
}

fn validates_with_reference_limit(limit: usize) -> TestResult<bool> {
    Ok(package_with_reference_limit(limit)?.validate().is_ok())
}

/// Find the smallest reference ceiling that admits the complete native
/// semantic projection.  Keeping the edge derived from the fixture makes the
/// test stable if unrelated native graph references are added later.
fn minimum_valid_reference_limit() -> TestResult<usize> {
    let maximum = SemanticLimits::MAX_REFERENCES;
    let mut lower = 1_usize;
    let mut upper = 1_usize;
    let mut upper_valid = validates_with_reference_limit(upper)?;

    for _ in 0..MAX_REFERENCE_SEARCH_STEPS {
        if upper_valid {
            break;
        }
        if upper == maximum {
            return Err(
                io::Error::other("native fixture exceeds the reference hard ceiling").into(),
            );
        }
        lower = upper
            .checked_add(1)
            .ok_or_else(|| io::Error::other("reference search overflowed"))?;
        upper = upper.saturating_mul(2).min(maximum);
        upper_valid = validates_with_reference_limit(upper)?;
    }
    if !upper_valid {
        return Err(io::Error::other("reference search exceeded its iteration bound").into());
    }

    for _ in 0..MAX_REFERENCE_SEARCH_STEPS {
        if lower >= upper {
            break;
        }
        let midpoint = lower + (upper - lower) / 2;
        if validates_with_reference_limit(midpoint)? {
            upper = midpoint;
        } else {
            lower = midpoint
                .checked_add(1)
                .ok_or_else(|| io::Error::other("reference search overflowed"))?;
        }
    }
    if lower != upper {
        return Err(io::Error::other("reference search exceeded its iteration bound").into());
    }
    Ok(lower)
}

#[test]
fn default_budget_allows_exact_native_audio_edit() -> TestResult {
    let audio = member_bytes(AUDIO_MEMBER)?;
    let mut replacement = audio.clone();
    let last = replacement
        .last_mut()
        .ok_or_else(|| io::Error::other("native audio member is empty"))?;
    *last ^= 1;

    let package = Package::from_bytes(NATIVE_SOURCE)?;
    package.validate()?;
    let commit = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?
        .set(&replacement)?
        .commit()?;

    assert!(!commit.patch().is_noop());
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )?,
        replacement.as_slice()
    );
    // The two native audio controls share one materialized data record.
    assert_eq!(
        commit.package().slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(1),
            MediaPart::Content,
        )?,
        replacement.as_slice()
    );
    Ok(())
}

#[test]
fn edge_reference_profile_rejects_media_operation_after_validation() -> TestResult {
    let reference_limit = minimum_valid_reference_limit()?;
    let package = package_with_reference_limit(reference_limit)?;
    package.validate()?;
    let source_before = exact_bytes(&package)?;
    assert_eq!(source_before, NATIVE_SOURCE);

    let audio = member_bytes(AUDIO_MEMBER)?;
    let mut replacement = audio.clone();
    let last = replacement
        .last_mut()
        .ok_or_else(|| io::Error::other("native audio member is empty"))?;
    *last ^= 1;

    let error = package
        .edit_slide_media_data(
            SlideSelector::index(0),
            MovieSelector::index(0),
            MediaPart::Content,
        )
        .and_then(|edit| edit.set(&replacement))
        .and_then(|edit| edit.commit())
        .err()
        .ok_or_else(|| {
            io::Error::other("the edge reference budget must reject the operation-local closure")
        })?;
    match error {
        SlideMediaDataError::LimitExceeded {
            kind: SlideMediaDataLimitKind::References,
            observed,
            maximum,
        } => assert!(observed > maximum),
        other => panic!("expected an aggregate reference limit, got {other:?}"),
    }
    assert_eq!(exact_bytes(&package)?, source_before);
    Ok(())
}
