#![cfg(feature = "iwork")]

use litchi::iwork::{Document, ErrorKind, Options, Resource, SourceLimits, Stage};

const PAGES: &[u8] = include_bytes!("../../../test-data/iwork/pages/basic.pages");

#[derive(Clone, Copy)]
enum Knob {
    Entries,
    EntryBytes,
    ExpandedBytes,
    DecodedBytes,
}

impl Knob {
    const fn resource(self) -> Resource {
        match self {
            Self::Entries => Resource::Entries,
            Self::EntryBytes => Resource::EntryBytes,
            Self::ExpandedBytes => Resource::AggregateDecodedBytes,
            Self::DecodedBytes => Resource::DecodedBytes,
        }
    }

    fn hard_maximum(self) -> u64 {
        match self {
            Self::Entries => u64::try_from(SourceLimits::HARD_MAX_ENTRIES).unwrap_or(u64::MAX),
            Self::EntryBytes => SourceLimits::HARD_MAX_ENTRY_BYTES,
            Self::ExpandedBytes => SourceLimits::HARD_MAX_EXPANDED_BYTES,
            Self::DecodedBytes => {
                u64::try_from(SourceLimits::HARD_MAX_DECODED_BYTES_PER_ITEM).unwrap_or(u64::MAX)
            },
        }
    }
}

fn options(knob: Knob, value: u64) -> Options {
    let defaults = SourceLimits::default();
    let entries = match knob {
        Knob::Entries => {
            usize::try_from(value).unwrap_or_else(|_error| panic!("entry limit must fit usize"))
        },
        _ => defaults.max_entries(),
    };
    let entry_bytes = match knob {
        Knob::EntryBytes => value,
        _ => defaults.max_entry_bytes(),
    };
    let expanded_bytes = match knob {
        Knob::ExpandedBytes => value,
        _ => defaults.max_expanded_bytes(),
    };
    let decoded_bytes = match knob {
        Knob::DecodedBytes => usize::try_from(value)
            .unwrap_or_else(|_error| panic!("decoded-byte limit must fit usize")),
        _ => defaults.max_decoded_bytes_per_item(),
    };
    let source = SourceLimits::new(
        defaults.max_input_bytes(),
        entries,
        entry_bytes,
        expanded_bytes,
        decoded_bytes,
    )
    .unwrap_or_else(|error| panic!("test profile must be valid: {error}"));
    Options::default().with_source(source)
}

fn accepts(knob: Knob, value: u64) -> bool {
    Document::from_bytes_with_options(PAGES, options(knob, value)).is_ok()
}

fn minimum_accepted(knob: Knob) -> u64 {
    let mut lower = 1;
    let mut upper = knob.hard_maximum();
    assert!(accepts(knob, upper), "hard maximum must accept the fixture");
    while lower < upper {
        let middle = lower + (upper - lower) / 2;
        if accepts(knob, middle) {
            upper = middle;
        } else {
            lower = middle + 1;
        }
    }
    lower
}

#[test]
fn physical_resources_accept_exact_limits_and_report_one_over() {
    for knob in [
        Knob::Entries,
        Knob::EntryBytes,
        Knob::ExpandedBytes,
        Knob::DecodedBytes,
    ] {
        let exact = minimum_accepted(knob);
        assert!(exact > 1, "native fixture must exercise a nontrivial limit");
        Document::from_bytes_with_options(PAGES, options(knob, exact))
            .unwrap_or_else(|error| panic!("exact physical limit must pass: {error}"));

        let maximum = exact - 1;
        let error = Document::from_bytes_with_options(PAGES, options(knob, maximum))
            .err()
            .unwrap_or_else(|| panic!("one-over physical resource must fail"));
        assert_eq!(error.kind(), ErrorKind::LimitExceeded);
        assert_eq!(error.stage(), Stage::Detection);
        assert_eq!(error.format(), None);
        assert_eq!(error.resource(), Some(knob.resource()));
        assert_eq!(error.observed(), Some(exact));
        assert_eq!(error.maximum(), Some(maximum));
        assert!(std::error::Error::source(&error).is_none());
    }
}

#[test]
fn input_bytes_accept_exact_limit_and_report_one_over_before_copying() {
    let defaults = SourceLimits::default();
    let exact =
        u64::try_from(PAGES.len()).unwrap_or_else(|_error| panic!("fixture length must fit u64"));
    let source = |maximum| {
        SourceLimits::new(
            maximum,
            defaults.max_entries(),
            defaults.max_entry_bytes(),
            defaults.max_expanded_bytes(),
            defaults.max_decoded_bytes_per_item(),
        )
        .unwrap_or_else(|error| panic!("test profile must be valid: {error}"))
    };
    Document::from_bytes_with_options(PAGES, Options::default().with_source(source(exact)))
        .unwrap_or_else(|error| panic!("exact input limit must pass: {error}"));

    let maximum = exact - 1;
    let error =
        Document::from_bytes_with_options(PAGES, Options::default().with_source(source(maximum)))
            .err()
            .unwrap_or_else(|| panic!("one-over input must fail"));
    assert_eq!(error.kind(), ErrorKind::LimitExceeded);
    assert_eq!(error.stage(), Stage::Input);
    assert_eq!(error.resource(), Some(Resource::InputBytes));
    assert_eq!(error.observed(), Some(exact));
    assert_eq!(error.maximum(), Some(maximum));
    assert!(std::error::Error::source(&error).is_none());
}
