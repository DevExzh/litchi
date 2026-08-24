#![no_main]

use std::fmt::{Debug, Display};
use std::hint::black_box;
use std::sync::OnceLock;

use libfuzzer_sys::fuzz_target;
use litchi::pages::{
    Limits, Package, SectionPaginationError, SectionPaginationLimitKind, SectionSelector,
    section::{PageNumber, PageNumbering, Pagination, Start},
};

const MAX_INPUT_BYTES: u64 = 256 * 1024;
const OVERSIZED_INPUT_BYTES: usize = 256 * 1024 + 1;
const MAX_ENTRIES: usize = 128;
const MAX_ENTRY_BYTES: u64 = 1024 * 1024;
const MAX_EXPANDED_BYTES: u64 = 4 * 1024 * 1024;
const MAX_IWA_STREAM_BYTES: usize = 1024 * 1024;
const PRIVATE_MALFORMED_INPUT: &[u8] = b"__litchi_private_pages_section_pagination_input_8d2e__";
const PRIVATE_SECTION_NAME: &str = "__litchi_private_pages_section_pagination_name_8d2e__";
const NATIVE_PAGES: &[u8] = include_bytes!("../../../../test-data/iwork/pages/basic.pages");

fuzz_target!(|data: &[u8]| {
    match Package::from_bytes_with_limits(data, fuzz_limits()) {
        Ok(package) => exercise_package(&package, data),
        Err(error) => observe_error(error),
    }

    // ZIP checksums make arbitrary bytes unlikely to reach the section codec.
    // Reuse every bounded input as a semantic command against the native
    // package so selector, transaction, inverse, and error paths stay hot.
    exercise_package(native_package(), data);
    exercise_semantic_values(data);
    exercise_resource_mutation(data);
    exercise_redacted_malformed_ingress();
    exercise_input_limit();
});

fn fuzz_limits() -> Limits {
    static LIMITS: OnceLock<Limits> = OnceLock::new();
    *LIMITS.get_or_init(|| {
        Limits::new(
            MAX_INPUT_BYTES,
            MAX_ENTRIES,
            MAX_ENTRY_BYTES,
            MAX_EXPANDED_BYTES,
            MAX_IWA_STREAM_BYTES,
        )
        .unwrap_or_else(|error| unreachable!("valid Pages pagination fuzz limits: {error}"))
    })
}

fn native_package() -> &'static Package {
    static PACKAGE: OnceLock<Package> = OnceLock::new();
    PACKAGE.get_or_init(|| {
        let package = Package::from_bytes_with_limits(NATIVE_PAGES, fuzz_limits())
            .unwrap_or_else(|error| panic!("native Pages pagination seed must open: {error}"));
        let expected = pagination(
            Some(Start::NextPage),
            Some(PageNumbering::ContinueFromPrevious),
            Some(1),
        );
        assert_eq!(
            package
                .section_pagination(SectionSelector::index(0))
                .unwrap_or_else(|error| {
                    panic!("native Pages pagination seed must expose pagination: {error}")
                }),
            expected
        );
        package
    })
}

fn exercise_package(package: &Package, data: &[u8]) {
    black_box((package.stats(), package.sections().len()));
    let source_before = package_bytes(package);

    observe_result(package.section_pagination(SectionSelector::index(0)));
    observe_result(package.section_pagination(SectionSelector::index(package.sections().len())));
    observe_result(package.section_pagination(SectionSelector::name("")));
    observe_result(package.section_pagination(SectionSelector::name("Blank")));
    if let Err(error) = package.section_pagination(SectionSelector::name(PRIVATE_SECTION_NAME)) {
        observe_redacted(error, PRIVATE_SECTION_NAME);
    }

    let before = match package.section_pagination(SectionSelector::index(0)) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };

    // Invalid aliases must not alter the staged value. These probes exercise
    // semantic validation independently from the physical rewrite below.
    let mut invalid = match package.edit_section_pagination(SectionSelector::index(0)) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, &source_before);
            return;
        },
    };
    if let Err(error) = invalid.set_start(Some(Start::Unknown(0))) {
        assert!(matches!(
            error,
            SectionPaginationError::InvalidPagination(_)
        ));
        observe_error(error);
    }
    if let Err(error) = invalid.set_page_numbering(Some(PageNumbering::Unknown(1))) {
        assert!(matches!(
            error,
            SectionPaginationError::InvalidPagination(_)
        ));
        observe_error(error);
    }
    assert_eq!(invalid.pagination(), before);
    drop(invalid);

    exercise_transaction(package, before, data, &source_before);
}

fn exercise_transaction(package: &Package, before: Pagination, data: &[u8], source_before: &[u8]) {
    let after = mutated_pagination(before, data);
    let mut edit = match package.edit_section_pagination(SectionSelector::index(0)) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source_before);
            return;
        },
    };

    let staged = match control(data, 0) % 7 {
        0 => edit.set_pagination(after).map(|_| ()),
        1 => edit.set_start(after.start()).map(|_| ()),
        2 => edit.set_page_numbering(after.page_numbering()).map(|_| ()),
        3 => {
            edit.set_starting_page_number(after.starting_page_number());
            Ok(())
        },
        4 => edit
            .set_start(after.start())
            .and_then(|edit| edit.set_page_numbering(after.page_numbering()))
            .map(|edit| {
                edit.set_starting_page_number(after.starting_page_number());
            }),
        5 => {
            edit.clear();
            Ok(())
        },
        _ => edit.set_pagination(after).map(|_| ()),
    };
    if let Err(error) = staged {
        observe_error(error);
        assert_source_unchanged(package, source_before);
        return;
    }

    // Reconstruct the value actually staged by the command so all setter
    // combinations are checked without retaining mutable editor state.
    let expected = match control(data, 0) % 7 {
        0 => after,
        1 => {
            let mut value = before;
            value
                .set_start(after.start())
                .unwrap_or_else(|error| unreachable!("canonical start: {error}"));
            value
        },
        2 => {
            let mut value = before;
            value
                .set_page_numbering(after.page_numbering())
                .unwrap_or_else(|error| unreachable!("canonical numbering: {error}"));
            value
        },
        3 => {
            let mut value = before;
            value.set_starting_page_number(after.starting_page_number());
            value
        },
        4 => {
            let mut value = before;
            value
                .set_start(after.start())
                .unwrap_or_else(|error| unreachable!("canonical start: {error}"));
            value
                .set_page_numbering(after.page_numbering())
                .unwrap_or_else(|error| unreachable!("canonical numbering: {error}"));
            value.set_starting_page_number(after.starting_page_number());
            value
        },
        5 => Pagination::new(),
        _ => after,
    };
    assert_eq!(edit.pagination(), expected);

    let commit = match edit.commit() {
        Ok(commit) => commit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(package, source_before);
            return;
        },
    };
    assert_source_unchanged(package, source_before);

    let patch = commit.patch().clone();
    let diagnostics = *commit.diagnostics();
    let changed = before != expected;
    assert_eq!(patch.before(), before);
    assert_eq!(patch.after(), expected);
    assert_eq!(patch.is_noop(), !changed);
    assert_eq!(diagnostics.changed(), changed);
    assert_eq!(diagnostics.touched_components(), usize::from(changed));
    assert_eq!(diagnostics.full_reparse_performed(), changed);
    assert_eq!(
        commit
            .package()
            .section_pagination(SectionSelector::index(0))
            .unwrap_or_else(|error| panic!("committed pagination must be readable: {error}")),
        expected
    );
    black_box((
        patch.position(),
        patch.source_fingerprint(),
        patch.target_fingerprint(),
        &patch,
        diagnostics,
    ));

    let applied = package
        .apply_section_pagination(&patch)
        .unwrap_or_else(|error| panic!("fresh pagination patch must apply: {error}"));
    assert_eq!(
        applied
            .package()
            .section_pagination(SectionSelector::index(0))
            .unwrap_or_else(|error| panic!("applied pagination must be readable: {error}")),
        expected
    );
    assert_eq!(
        package_bytes(applied.package()),
        package_bytes(commit.package())
    );

    let inverse = patch.inverse();
    assert_eq!(inverse.inverse(), patch);
    if changed {
        assert!(matches!(
            applied.package().apply_section_pagination(&patch),
            Err(SectionPaginationError::PatchConflict)
        ));
        assert!(matches!(
            package.apply_section_pagination(&inverse),
            Err(SectionPaginationError::PatchConflict)
        ));
    }
    let restored = applied
        .package()
        .apply_section_pagination(&inverse)
        .unwrap_or_else(|error| panic!("pagination inverse must apply: {error}"));
    assert_eq!(package_bytes(restored.package()), source_before);
    assert_eq!(
        restored
            .package()
            .section_pagination(SectionSelector::index(0))
            .unwrap_or_else(|error| panic!("restored pagination must be readable: {error}")),
        before
    );
    assert_source_unchanged(package, source_before);
}

fn exercise_resource_mutation(data: &[u8]) {
    // Keep ingress just above the native source while varying the same budget
    // that also caps physical output. A max-valued pagination forces either a
    // bounded output rejection or a complete candidate that still inverts.
    let source_size = u64::try_from(NATIVE_PAGES.len())
        .unwrap_or_else(|error| unreachable!("native Pages size fits u64: {error}"));
    let budget = source_size.saturating_add(u64::from(control(data, 20) & 7));
    let limits = match Limits::new(
        budget,
        MAX_ENTRIES,
        MAX_ENTRY_BYTES,
        MAX_EXPANDED_BYTES,
        MAX_IWA_STREAM_BYTES,
    ) {
        Ok(limits) => limits,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let package = match Package::from_bytes_with_limits(NATIVE_PAGES, limits) {
        Ok(package) => package,
        Err(error) => {
            observe_error(error);
            return;
        },
    };
    let source_before = package_bytes(&package);
    let before = match package.section_pagination(SectionSelector::index(0)) {
        Ok(value) => value,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(&package, &source_before);
            return;
        },
    };
    let target = maximal_pagination();
    let mut edit = match package.edit_section_pagination(SectionSelector::index(0)) {
        Ok(edit) => edit,
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(&package, &source_before);
            return;
        },
    };
    edit.set_pagination(target)
        .unwrap_or_else(|error| panic!("maximal canonical pagination is valid: {error}"));
    match edit.commit() {
        Ok(commit) => {
            assert_source_unchanged(&package, &source_before);
            let inverse = commit.patch().inverse();
            let restored = commit
                .package()
                .apply_section_pagination(&inverse)
                .unwrap_or_else(|error| panic!("resource-profile inverse must apply: {error}"));
            assert_eq!(package_bytes(restored.package()), source_before);
            assert_eq!(
                restored
                    .package()
                    .section_pagination(SectionSelector::index(0))
                    .unwrap_or_else(|error| {
                        panic!("resource-profile restored pagination must be readable: {error}")
                    }),
                before
            );
        },
        Err(error) => {
            observe_error(error);
            assert_source_unchanged(&package, &source_before);
        },
    }
}

fn exercise_semantic_values(data: &[u8]) {
    let raw = read_u32(data, 4);
    observe_result(Start::unknown(raw));
    observe_result(PageNumbering::unknown(raw));
    observe_result(PageNumber::new(raw));

    let mut invalid = Pagination::new();
    observe_result(invalid.set_start(Some(Start::Unknown(0))));
    observe_result(invalid.set_page_numbering(Some(PageNumbering::Unknown(1))));
    observe_result(invalid.validate());
    black_box(invalid);

    // Keep the error vocabulary live even when the package ingress rejects
    // the current mutation before reaching a section transaction.
    black_box(SectionPaginationLimitKind::OutputBytes.to_string());
}

fn exercise_redacted_malformed_ingress() {
    match Package::from_bytes_with_limits(PRIVATE_MALFORMED_INPUT, fuzz_limits()) {
        Err(error) => observe_redacted(
            error,
            std::str::from_utf8(PRIVATE_MALFORMED_INPUT)
                .unwrap_or_else(|error| unreachable!("private sentinel is UTF-8: {error}")),
        ),
        Ok(_) => panic!("a private malformed Pages sentinel must not parse"),
    }
}

fn exercise_input_limit() {
    static OVERSIZED: OnceLock<Box<[u8]>> = OnceLock::new();
    let bytes = OVERSIZED.get_or_init(|| vec![0; OVERSIZED_INPUT_BYTES].into_boxed_slice());
    match Package::from_bytes_with_limits(bytes, fuzz_limits()) {
        Err(error) => observe_error(error),
        Ok(_) => panic!("an oversized Pages input must be rejected"),
    }
}

fn mutated_pagination(before: Pagination, data: &[u8]) -> Pagination {
    let mut after = before;
    match control(data, 0) % 7 {
        0 => {},
        1 => {
            after
                .set_start(distinct_start(before.start(), data))
                .unwrap_or_else(|error| unreachable!("canonical start is valid: {error}"));
        },
        2 => {
            after
                .set_page_numbering(distinct_numbering(before.page_numbering(), data))
                .unwrap_or_else(|error| unreachable!("canonical numbering is valid: {error}"));
        },
        3 => after.set_starting_page_number(distinct_page(before.starting_page_number(), data)),
        4 => {
            after
                .set_start(distinct_start(before.start(), data))
                .unwrap_or_else(|error| unreachable!("canonical start is valid: {error}"));
            after
                .set_page_numbering(distinct_numbering(before.page_numbering(), data))
                .unwrap_or_else(|error| unreachable!("canonical numbering is valid: {error}"));
            after.set_starting_page_number(distinct_page(before.starting_page_number(), data));
        },
        5 => after = Pagination::new(),
        _ => after = maximal_pagination(),
    }
    after
}

fn distinct_start(current: Option<Start>, data: &[u8]) -> Option<Start> {
    let raw = read_u32(data, 8).max(3);
    let values = [
        None,
        Some(Start::NextPage),
        Some(Start::RightPage),
        Some(Start::LeftPage),
        Some(Start::Unknown(raw)),
    ];
    choose_distinct(current, &values, data, 1)
}

fn distinct_numbering(current: Option<PageNumbering>, data: &[u8]) -> Option<PageNumbering> {
    let raw = read_u32(data, 12).max(2);
    let values = [
        None,
        Some(PageNumbering::ContinueFromPrevious),
        Some(PageNumbering::Restart),
        Some(PageNumbering::Unknown(raw)),
    ];
    choose_distinct(current, &values, data, 2)
}

fn distinct_page(current: Option<PageNumber>, data: &[u8]) -> Option<PageNumber> {
    let raw = read_u32(data, 16).max(1);
    let values = [
        None,
        Some(PageNumber::new(1).unwrap_or_else(|error| unreachable!("one is nonzero: {error}"))),
        Some(PageNumber::new(42).unwrap_or_else(|error| unreachable!("42 is nonzero: {error}"))),
        Some(
            PageNumber::new(u32::MAX)
                .unwrap_or_else(|error| unreachable!("u32::MAX is nonzero: {error}")),
        ),
        Some(PageNumber::new(raw).unwrap_or_else(|error| unreachable!("bounded page: {error}"))),
    ];
    choose_distinct(current, &values, data, 3)
}

fn choose_distinct<T: Copy + PartialEq>(
    current: Option<T>,
    values: &[Option<T>],
    data: &[u8],
    offset: usize,
) -> Option<T> {
    let index = usize::from(control(data, offset)) % values.len();
    let selected = values[index];
    if selected != current {
        selected
    } else {
        values[(index + 1) % values.len()]
    }
}

fn maximal_pagination() -> Pagination {
    let mut value = Pagination::new();
    value
        .set_start(Some(Start::Unknown(u32::MAX)))
        .unwrap_or_else(|error| unreachable!("unknown start is canonical: {error}"));
    value
        .set_page_numbering(Some(PageNumbering::Unknown(u32::MAX)))
        .unwrap_or_else(|error| unreachable!("unknown numbering is canonical: {error}"));
    value.set_starting_page_number(Some(
        PageNumber::new(u32::MAX)
            .unwrap_or_else(|error| unreachable!("u32::MAX is nonzero: {error}")),
    ));
    value
}

fn pagination(
    start: Option<Start>,
    numbering: Option<PageNumbering>,
    page: Option<u32>,
) -> Pagination {
    let mut value = Pagination::new();
    value
        .set_start(start)
        .unwrap_or_else(|error| unreachable!("native start is canonical: {error}"));
    value
        .set_page_numbering(numbering)
        .unwrap_or_else(|error| unreachable!("native numbering is canonical: {error}"));
    value.set_starting_page_number(page.map(|number| {
        PageNumber::new(number)
            .unwrap_or_else(|error| unreachable!("native page is nonzero: {error}"))
    }));
    value
}

fn assert_source_unchanged(package: &Package, source_before: &[u8]) {
    assert_eq!(package_bytes(package), source_before);
}

fn read_u32(data: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes([
        control(data, offset),
        control(data, offset + 1),
        control(data, offset + 2),
        control(data, offset + 3),
    ])
}

fn control(data: &[u8], index: usize) -> u8 {
    data.get(index).copied().unwrap_or_default()
}

fn package_bytes(package: &Package) -> Vec<u8> {
    let mut bytes = Vec::new();
    package
        .write_to(&mut bytes)
        .unwrap_or_else(|error| panic!("writing a Pages package to memory must succeed: {error}"));
    bytes
}

fn observe_result<T, E>(result: Result<T, E>)
where
    T: Debug,
    E: Debug + Display,
{
    match result {
        Ok(value) => {
            black_box(value);
        },
        Err(error) => observe_error(error),
    }
}

fn observe_error(error: impl Debug + Display) {
    black_box(error.to_string());
    black_box(format!("{error:?}"));
}

fn observe_redacted(error: impl Debug + Display, private: &str) {
    let display = error.to_string();
    let debug = format!("{error:?}");
    assert!(!display.contains(private));
    assert!(!debug.contains(private));
    black_box((display, debug));
}
