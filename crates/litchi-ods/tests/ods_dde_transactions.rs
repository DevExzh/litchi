//! Integration fixtures for the inert ODS DDE transaction surface.
//!
//! The source contains two ordered formula links, two worksheet-local sources,
//! repeated cached values, and opaque markup that a transaction must either
//! retain byte-for-byte or reject before staging a destructive replacement.
//! Every source uses a deliberately nonexistent application/topic and disables
//! automatic updates; these tests never resolve or refresh an external target.

use std::{
    fmt::Write as _,
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits, Profile,
    Resource,
};
use litchi_ods::dde::{
    self, AutomaticUpdate, CachedCell, CachedRow, CachedTable, CachedValue, ConversionMode,
    LinkSpec,
};

/// A compact synthetic ODF 1.4 content owner used by the transaction cases.
///
/// `office:name="Shared"` intentionally occurs on both links.  A link is
/// selected by source order in the transaction tests; a name-only selector
/// must report ambiguity rather than silently selecting one of them.
const CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?><?producer before?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:vendor="urn:example:opaque-dde" office:version="1.4"><office:body><office:spreadsheet><?spreadsheet before?><table:table table:name="Data"><office:dde-source office:dde-application="never-contacted-dde" office:dde-topic="file:///never-contacted-dde.ods" office:dde-item="Data.A1:B2" office:name="DataSource" office:conversion-mode="keep-text" office:automatic-update="false"/><table:table-row><table:table-cell office:value-type="string"><text:p>untouched</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Second"><office:dde-source office:dde-application="never-contacted-dde" office:dde-topic="file:///never-contacted-dde.ods" office:dde-item="Second.A1" office:name="SecondSource" office:conversion-mode="into-english-number" office:automatic-update="false"/><table:table-row><table:table-cell office:value-type="string"><text:p>second</text:p></table:table-cell></table:table-row></table:table><?spreadsheet before links?><table:dde-links><!-- keep link-owner comment --><table:dde-link><?link-one?><office:dde-source office:dde-application="never-contacted-dde" office:dde-topic="file:///never-contacted-dde.ods" office:dde-item="Data.A1:B2" office:name="Shared" office:conversion-mode="keep-text" office:automatic-update="false"/><table:table table:name="CacheOne"><table:table-column table:number-columns-repeated="3"/><!-- opaque cache comment --><?cache-opaque keep?><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="2"/><table:table-cell office:value-type="string" office:string-value="repeat-tail"/></table:table-row></table:table></table:dde-link><table:dde-link><office:dde-source office:dde-application="never-contacted-dde" office:dde-topic="file:///never-contacted-dde.ods" office:dde-item="Second.C3" office:name="Shared" office:conversion-mode="into-english-number" office:automatic-update="false"/><table:table table:name="CacheTwo"><table:table-column/><table:table-row><table:table-cell office:value-type="string" office:string-value="second-cache"/><table:table-cell office:value-type="string" office:string-value="second-cache"/></table:table-row></table:table></table:dde-link></table:dde-links><vendor:spreadsheet-extension vendor:keep="yes"><vendor:value>untouched sibling</vendor:value></vendor:spreadsheet-extension><?producer after?></office:spreadsheet></office:body></office:document-content>"#;

fn many_links_content(link_count: usize) -> String {
    let mut content = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table><table:dde-links>"#,
    );
    for index in 0..link_count {
        write!(
            content,
            r#"<table:dde-link><office:dde-source office:dde-application="never-contacted-dde" office:dde-topic="file:///never-contacted-dde.ods" office:dde-item="Data.A1" office:name="Link{index}" office:conversion-mode="keep-text" office:automatic-update="false"/><table:table table:name="Cache{index}"><table:table-column/><table:table-row><table:table-cell office:value-type="string" office:string-value="value{index}"/></table:table-row></table:table></table:dde-link>"#,
            index = index
        )
        .expect("many-link fixture formatting");
    }
    content.push_str(
        r#"</table:dde-links></office:spreadsheet></office:body></office:document-content>"#,
    );
    content
}

fn snapshot() -> dde::Snapshot {
    dde::Snapshot::parse(CONTENT).expect("synthetic DDE fixture should parse")
}

fn context(scope: &str) -> ExecutionContext {
    managed_context(scope).2
}

fn managed_context(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), BudgetLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(1024 * 1024).expect("nonzero byte count"),
        0,
    )
    .expect("finite execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn bounded_context(
    scope: &str,
    memory_bytes: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        scope.to_owned(),
        BudgetLimits::new(
            memory_bytes,
            2 * 1024 * 1024 * 1024,
            4 * 1024 * 1024 * 1024,
            10_000_000,
            256,
            1_000_000_000,
        ),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(1024 * 1024).expect("nonzero byte count"),
        0,
    )
    .expect("finite execution limits");
    (
        budget.clone(),
        cancellation,
        ExecutionContext::new(budget, token, limits),
    )
}

fn typed_source(name: &str, item: &str) -> dde::Source {
    dde::Source::new(
        "never-contacted-dde",
        "file:///never-contacted-dde.ods",
        item,
    )
    .expect("valid inert source")
    .named(name)
    .expect("valid source name")
    .with_conversion_mode(ConversionMode::KeepText)
    .with_automatic_update(AutomaticUpdate::Disabled)
}

fn typed_cache() -> CachedTable {
    let row = CachedRow::from_cells(vec![
        CachedCell::new(CachedValue::Number(7.0)).expect("finite number"),
        CachedCell::new(CachedValue::Number(7.0)).expect("repeated finite number"),
        CachedCell::new(CachedValue::Text("new cache".to_owned())).expect("cache text"),
        CachedCell::new(CachedValue::Boolean(true)).expect("cache bool"),
    ])
    .expect("bounded cache row");
    CachedTable::from_rows(vec![row]).expect("bounded cache table")
}

fn typed_link(name: &str, item: &str) -> LinkSpec {
    LinkSpec::new(typed_source(name, item), typed_cache()).expect("complete typed DDE link")
}

fn large_typed_link(name: &str, item: &str, cell_count: usize, text: &str) -> LinkSpec {
    let cell = CachedCell::new(CachedValue::Text(text.to_owned())).expect("large cache text");
    let row = CachedRow::from_cells(vec![cell; cell_count]).expect("large cache row");
    let cache = CachedTable::from_rows(vec![row]).expect("large cache table");
    LinkSpec::new(typed_source(name, item), cache).expect("complete large typed DDE link")
}

fn long_source_only_source() -> dde::Source {
    let name = "source-only-long-name-".repeat(2_800);
    typed_source(&name, "Data.XFD1")
}

fn assert_sheet_source_payload_refused(
    content: &str,
    scope: &str,
    operation: impl FnOnce(&mut dde::Edit, dde::Source) -> litchi_core::Result<()>,
) {
    let (budget, _cancellation, execution) = bounded_context(scope, 64 * 1024);
    let source = dde::Snapshot::parse_with_context(content, dde::Limits::default(), &execution)
        .expect("the small source snapshot should fit the tight profile");
    let baseline_memory = budget.used(Resource::Memory);
    let before = source.source_xml().to_owned();
    let original_source_count = source.sheet_sources().len();
    let mut edit = source.edit();
    let error = match operation(&mut edit, long_source_only_source()) {
        Err(error) => error,
        Ok(()) => panic!("the long sheet-source payload must exceed the tight memory profile"),
    };
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("memory"),
        "sheet-source refusal should identify memory admission: {error}"
    );
    assert!(
        edit.is_noop(),
        "failed sheet-source admission must leave the edit unchanged"
    );
    assert_eq!(edit.sheet_source_count(), original_source_count);
    assert_eq!(edit.before().source_xml(), before);
    assert_eq!(
        budget.used(Resource::Memory),
        baseline_memory,
        "a refused sheet-source payload must release every temporary reservation"
    );
    drop(edit);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn content_without_second_sheet_source() -> String {
    clean_link_owner_content().replace(
        "<office:dde-source office:dde-application=\"never-contacted-dde\" office:dde-topic=\"file:///never-contacted-dde.ods\" office:dde-item=\"Second.A1\" office:name=\"SecondSource\" office:conversion-mode=\"into-english-number\" office:automatic-update=\"false\"/>",
        "",
    )
}

fn clean_link_owner_content() -> String {
    CONTENT.replace("<!-- keep link-owner comment -->", "")
}

fn typed_cache_content() -> String {
    clean_link_owner_content()
        .replace("<?link-one?>", "")
        .replace("<!-- opaque cache comment -->", "")
        .replace("<?cache-opaque keep?>", "")
}

fn unnamed_typed_cache_content() -> String {
    typed_cache_content().replace(" table:name=\"CacheOne\"", "")
}

#[test]
fn fixture_exposes_ordered_sources_and_exact_repeated_cache_xml() {
    let snapshot = snapshot();

    assert_eq!(snapshot.source_xml(), CONTENT);
    assert_eq!(snapshot.sheet_sources().len(), 2);
    assert_eq!(snapshot.sheet_sources()[0].sheet(), "Data");
    assert_eq!(snapshot.sheet_sources()[1].sheet(), "Second");
    assert_eq!(
        snapshot.sheet_sources()[0].source().automatic_update(),
        AutomaticUpdate::Disabled
    );
    assert_eq!(
        snapshot.sheet_sources()[1].source().conversion_mode(),
        ConversionMode::IntoEnglishNumber
    );

    assert_eq!(snapshot.links().len(), 2);
    assert_eq!(snapshot.links()[0].source().name(), Some("Shared"));
    assert_eq!(snapshot.links()[1].source().name(), Some("Shared"));
    assert_eq!(snapshot.links()[0].source().item(), "Data.A1:B2");
    assert_eq!(snapshot.links()[1].source().item(), "Second.C3");
    assert_eq!(
        snapshot.links()[0].cached_table_xml(),
        r#"<table:table table:name="CacheOne"><table:table-column table:number-columns-repeated="3"/><!-- opaque cache comment --><?cache-opaque keep?><table:table-row table:number-rows-repeated="2"><table:table-cell office:value-type="float" office:value="7" table:number-columns-repeated="2"/><table:table-cell office:value-type="string" office:string-value="repeat-tail"/></table:table-row></table:table>"#
    );
    assert!(
        snapshot.links()[0]
            .cached_table_xml()
            .contains("number-rows-repeated=\"2\"")
    );
    assert!(
        snapshot.links()[0]
            .cached_table_xml()
            .contains("number-columns-repeated=\"2\"")
    );
    assert!(
        snapshot.links()[0]
            .cached_table_xml()
            .contains("opaque cache comment")
    );
}

#[test]
fn fixture_keeps_external_targets_inert_and_rejects_basic_source_errors() {
    let snapshot = snapshot();
    for link in snapshot.links() {
        assert_eq!(link.source().application(), "never-contacted-dde");
        assert_eq!(link.source().automatic_update(), AutomaticUpdate::Disabled);
    }

    assert!(dde::Source::new("", "topic", "item").is_err());
    assert!(dde::Source::new("app", "topic", "item\u{0}").is_err());
    let limits = dde::Limits::default().with_input_bytes(CONTENT.len() - 1);
    assert!(dde::Snapshot::parse_with(CONTENT, limits).is_err());
}

#[test]
fn no_op_commit_is_exact_and_link_reorder_has_an_exact_inverse() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("clean DDE fixture");

    let mut no_op = source.edit();
    let no_op_commit = no_op
        .commit(&context("ods-dde-no-op"))
        .expect("no-op commit");
    assert!(!no_op_commit.changed());
    assert!(no_op_commit.patch().is_empty());
    assert_eq!(no_op_commit.snapshot().source_xml(), content);

    let mut reorder = source.edit();
    reorder
        .move_link(0usize, 1usize)
        .expect("move second link first");
    let reordered = reorder
        .commit(&context("ods-dde-reorder"))
        .expect("reordered links");
    assert!(reordered.changed());
    assert_eq!(reordered.snapshot().links()[0].source().item(), "Second.C3");
    assert_eq!(
        reordered.snapshot().links()[1].source().item(),
        "Data.A1:B2"
    );

    let restored = reordered
        .patch()
        .inverse()
        .apply(reordered.snapshot())
        .expect("inverse reorder patch");
    assert_eq!(restored.snapshot().source_xml(), content);

    let conflict_content = format!("{content} ");
    let conflict = dde::Snapshot::parse(&conflict_content).expect("conflicting XML remains valid");
    assert!(
        reordered.patch().apply(&conflict).is_err(),
        "patches require exact source bytes"
    );
}

#[test]
fn duplicate_link_names_and_missing_sheet_selectors_are_refused() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("clean DDE fixture");
    let mut edit = source.edit();

    assert!(
        edit.replace_link_source_named("Shared", typed_source("Ambiguous", "Data.X1"))
            .is_err()
    );
    assert!(
        edit.set_sheet_source("Missing", typed_source("Missing", "Data.X1"))
            .is_err()
    );
    assert!(
        edit.set_sheet_source(99usize, typed_source("Missing", "Data.X1"))
            .is_err()
    );
    assert!(edit.is_noop(), "refused selectors must not stage changes");
}

#[test]
fn duplicate_sheet_names_require_position_for_source_edits() {
    let content = CONTENT.replace(
        "<table:table table:name=\"Second\">",
        "<table:table table:name=\"Data\">",
    );
    assert_ne!(
        content, CONTENT,
        "duplicate-name fixture replacement applied"
    );
    let source = dde::Snapshot::parse(&content).expect("duplicate-name fixture");
    assert_eq!(source.sheet_sources().len(), 2);
    assert_eq!(source.sheet_sources()[0].sheet(), "Data");
    assert_eq!(source.sheet_sources()[1].sheet(), "Data");

    let mut by_name = source.edit();
    assert!(
        by_name
            .set_sheet_source("Data", typed_source("Ambiguous", "Data.X1"))
            .is_err(),
        "duplicate worksheet names must not select an arbitrary table"
    );
    assert!(by_name.is_noop());

    let mut by_position = source.edit();
    by_position
        .set_sheet_source(1usize, typed_source("SecondByPosition", "Second.Z9"))
        .expect("physical worksheet position selects the second duplicate");
    let committed = by_position
        .commit(&context("ods-dde-duplicate-sheet-position"))
        .expect("commit positional source edit");
    let sources = committed.snapshot().sheet_sources();
    assert_eq!(sources.len(), 2);
    assert_eq!(sources[0].sheet(), "Data");
    assert_eq!(sources[0].source().name(), Some("DataSource"));
    assert_eq!(sources[0].source().item(), "Data.A1:B2");
    assert_eq!(sources[1].sheet(), "Data");
    assert_eq!(sources[1].source().name(), Some("SecondByPosition"));
    assert_eq!(sources[1].source().item(), "Second.Z9");
}

#[test]
fn flat_spreadsheet_root_supports_source_edit_and_inverse() {
    let flat = clean_link_owner_content()
        .replace("office:document-content", "office:document")
        .replace(
            "office:version=\"1.4\"",
            "office:mimetype=\"application/vnd.oasis.opendocument.spreadsheet\"",
        );
    let source = dde::Snapshot::parse(&flat).expect("flat ODS root fixture");
    let mut edit = source.edit();
    edit.replace_link_source(0usize, typed_source("FlatEdited", "Data.F4"))
        .expect("stage flat-root source edit");
    let committed = edit
        .commit(&context("ods-dde-flat-root-edit"))
        .expect("commit flat-root source edit");
    assert!(committed.changed());
    assert_eq!(
        committed.snapshot().links()[0].source().name(),
        Some("FlatEdited")
    );
    assert_eq!(committed.snapshot().links()[0].source().item(), "Data.F4");

    let restored = committed
        .patch()
        .inverse()
        .apply(committed.snapshot())
        .expect("apply flat-root inverse patch");
    assert_eq!(restored.snapshot().source_xml(), flat);
}

#[test]
fn removing_the_last_link_and_all_links_is_a_changed_commit() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("clean DDE fixture");

    let mut remove_last = source.edit();
    remove_last.remove_link(1usize).expect("remove final link");
    let after_last = remove_last
        .commit(&context("ods-dde-remove-last"))
        .expect("commit final-link removal");
    assert!(after_last.changed());
    assert_eq!(after_last.snapshot().links().len(), 1);
    assert!(
        after_last
            .snapshot()
            .source_xml()
            .contains("table:dde-links")
    );

    let mut remove_all = source.edit();
    remove_all.remove_link(1usize).expect("remove second link");
    remove_all.remove_link(0usize).expect("remove first link");
    assert!(!remove_all.is_noop(), "link length change is never a no-op");
    let after_all = remove_all
        .commit(&context("ods-dde-remove-all"))
        .expect("commit all-link removal");
    assert!(after_all.changed());
    assert!(after_all.snapshot().links().is_empty());
    assert!(
        !after_all
            .snapshot()
            .source_xml()
            .contains("table:dde-links")
    );
}

#[test]
fn sheet_source_crud_preserves_order_and_readback() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("clean DDE fixture");
    let replacement = typed_source("DataSource2", "Data.C5");

    let mut replace = source.edit();
    replace
        .set_sheet_source("Data", replacement.clone())
        .expect("replace Data source");
    let replaced = replace
        .commit(&context("ods-dde-sheet-replace"))
        .expect("commit source replacement");
    assert_eq!(replaced.snapshot().sheet_sources().len(), 2);
    assert_eq!(
        replaced.snapshot().sheet_sources()[0].source().item(),
        "Data.C5"
    );

    let mut remove = replaced.snapshot().edit();
    let removed = remove
        .remove_sheet_source("Data")
        .expect("remove Data source");
    assert_eq!(removed.item(), "Data.C5");
    let removed_commit = remove
        .commit(&context("ods-dde-sheet-remove"))
        .expect("commit source removal");
    assert_eq!(removed_commit.snapshot().sheet_sources().len(), 1);
    assert_eq!(
        removed_commit.snapshot().sheet_sources()[0].sheet(),
        "Second"
    );

    let mut add = removed_commit.snapshot().edit();
    add.set_sheet_source("Data", replacement)
        .expect("add Data source back");
    let added = add
        .commit(&context("ods-dde-sheet-add"))
        .expect("commit source add");
    assert_eq!(added.snapshot().sheet_sources().len(), 2);
    assert_eq!(added.snapshot().sheet_sources()[0].sheet(), "Data");
    assert_eq!(
        added.snapshot().sheet_sources()[0].source().name(),
        Some("DataSource2")
    );
}

#[test]
fn set_sheet_source_payload_refusal_is_atomic_under_finite_memory() {
    assert_sheet_source_payload_refused(
        &clean_link_owner_content(),
        "ods-dde-sheet-set-memory-refusal",
        |edit, source| edit.set_sheet_source("Data", source),
    );
}

#[test]
fn replace_sheet_source_payload_refusal_is_atomic_under_finite_memory() {
    assert_sheet_source_payload_refused(
        &clean_link_owner_content(),
        "ods-dde-sheet-replace-memory-refusal",
        |edit, source| edit.replace_sheet_source("Data", source),
    );
}

#[test]
fn removing_an_absent_sheet_source_does_not_initialize_drafts() {
    let content = content_without_second_sheet_source();
    assert!(!content.contains("office:name=\"SecondSource\""));
    let (budget, _cancellation, execution) =
        bounded_context("ods-dde-sheet-remove-absent", 128 * 1024);
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &execution)
        .expect("fixture with one sheet source should fit the finite profile");
    assert_eq!(source.sheet_sources().len(), 1);
    assert_eq!(source.sheet_sources()[0].sheet(), "Data");
    let baseline_memory = budget.used(Resource::Memory);
    let before = source.source_xml().to_owned();
    let mut edit = source.edit();
    let error = edit
        .remove_sheet_source("Second")
        .expect_err("an absent source must be refused");
    assert!(error.to_string().contains("source was not found"));
    assert!(edit.is_noop());
    assert_eq!(edit.sheet_source_count(), 1);
    assert_eq!(edit.before().source_xml(), before);
    assert_eq!(budget.used(Resource::Memory), baseline_memory);
    drop(edit);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn named_plain_cache_replacement_preserves_table_name_and_inverts_exactly() {
    let content = typed_cache_content();
    let source = dde::Snapshot::parse(&content).expect("plain named-cache fixture");
    assert!(
        !source.links()[0]
            .cached_table_xml()
            .contains("opaque cache"),
        "fixture must contain no opaque cache markup"
    );
    assert!(
        !source.links()[0]
            .cached_table_xml()
            .contains("cache-opaque"),
        "fixture must contain no opaque cache processing instruction"
    );
    assert!(
        source.links()[0]
            .cached_table_xml()
            .contains("table:name=\"CacheOne\""),
        "the regression depends on a plain cache with a name"
    );

    let mut edit = source.edit();
    edit.replace_link_at(0usize, typed_link("NamedReplacement", "Data.P1"))
        .expect("stage named plain-cache replacement");
    let committed = edit
        .commit(&context("ods-dde-named-plain-cache-replacement"))
        .expect("replace named plain cache");
    assert!(committed.changed());
    let replacement_cache = committed.snapshot().links()[0].cached_table_xml();
    assert!(replacement_cache.contains("table:name=\"CacheOne\""));
    assert!(replacement_cache.contains("office:string-value=\"new cache\""));
    assert_eq!(
        committed.snapshot().links()[0].source().name(),
        Some("NamedReplacement")
    );
    assert_eq!(committed.snapshot().links()[0].source().item(), "Data.P1");

    let restored = committed
        .patch()
        .inverse()
        .apply(committed.snapshot())
        .expect("inverse named plain-cache replacement");
    assert_eq!(restored.snapshot().source_xml(), content);
}

fn assert_typed_cache_replacement_refused(content: String, scope: &str) {
    let source = dde::Snapshot::parse(&content).expect("cache refusal fixture");
    let before = source.source_xml().to_owned();
    let mut edit = source.edit();
    edit.replace_link_at(0usize, typed_link("ShouldRefuse", "Data.R1"))
        .expect("stage typed replacement for cache refusal");
    let error = match edit.commit(&context(scope)) {
        Err(error) => error,
        Ok(_) => panic!("unmodeled cache markup must refuse typed replacement"),
    };
    assert!(
        error.to_string().contains("unknown markup"),
        "unexpected refusal: {error}"
    );
    assert_eq!(edit.before().source_xml(), before);
}

#[test]
fn typed_cache_replacement_refuses_unmodeled_style_attributes() {
    let content = unnamed_typed_cache_content().replace(
        "<table:table-cell office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>",
        "<table:table-cell table:style-name=\"UnmodeledStyle\" office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>",
    );
    assert_typed_cache_replacement_refused(content, "ods-dde-style-cache-refusal");
}

#[test]
fn typed_cache_replacement_refuses_covered_table_cells() {
    let content = unnamed_typed_cache_content().replace(
        "<table:table-cell office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>",
        "<table:covered-table-cell/>",
    );
    assert_typed_cache_replacement_refused(content, "ods-dde-covered-cache-refusal");
}

#[test]
fn typed_cache_replacement_refuses_unmodeled_text_spans() {
    let content = unnamed_typed_cache_content().replace(
        "<table:table-cell office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>",
        "<table:table-cell office:value-type=\"string\"><text:p><text:span/></text:p></table:table-cell>",
    );
    assert_typed_cache_replacement_refused(content, "ods-dde-span-cache-refusal");
}

#[test]
fn typed_cache_replacement_refuses_unmodeled_cell_span_attributes() {
    let needle = "<table:table-cell office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>";
    for (index, attribute) in [
        "table:number-columns-spanned=\"2\"",
        "table:number-rows-spanned=\"2\"",
        "table:number-matrix-columns-spanned=\"2\"",
        "table:number-matrix-rows-spanned=\"2\"",
    ]
    .into_iter()
    .enumerate()
    {
        let replacement = format!("<table:table-cell {attribute}/>");
        let content = unnamed_typed_cache_content().replace(needle, &replacement);
        assert_typed_cache_replacement_refused(
            content,
            &format!("ods-dde-cell-span-cache-refusal-{index}"),
        );
    }
}

#[test]
fn typed_link_crud_emits_repeated_scalar_values() {
    let content = typed_cache_content();
    let source = dde::Snapshot::parse(&content).expect("typed-cache fixture");
    let mut add = source.edit();
    add.add_link(typed_link("Added", "Data.D4"))
        .expect("append typed link");
    let added = add
        .commit(&context("ods-dde-link-add"))
        .expect("commit typed link add");
    assert_eq!(added.snapshot().links().len(), 3);
    let added_cache = added.snapshot().links()[2].cached_table_xml();
    assert!(added_cache.contains("office:value-type=\"float\""));
    assert!(added_cache.contains("office:value-type=\"boolean\""));
    assert_eq!(added_cache.matches("office:value=\"7\"").count(), 2);
    assert!(added_cache.contains("office:string-value=\"new cache\""));
    assert_eq!(
        added_cache
            .matches("office:string-value=\"new cache\"")
            .count(),
        1
    );
    assert!(
        !added_cache.contains("<text:p>"),
        "authored DDE cache cells carry scalar attributes only"
    );

    let mut replace = added.snapshot().edit();
    replace
        .replace_link_at(2usize, typed_link("Replaced", "Data.E5"))
        .expect("replace typed link");
    let replaced = replace
        .commit(&context("ods-dde-link-replace"))
        .expect("commit typed link replacement");
    assert_eq!(
        replaced.snapshot().links()[2].source().name(),
        Some("Replaced")
    );

    let mut remove = replaced.snapshot().edit();
    remove.remove_link(2usize).expect("remove typed link");
    let removed = remove
        .commit(&context("ods-dde-link-remove"))
        .expect("commit typed link removal");
    assert_eq!(removed.snapshot().links().len(), 2);
}

#[test]
fn empty_authored_cache_emits_a_schema_valid_cell() {
    let content = typed_cache_content();
    let source = dde::Snapshot::parse(&content).expect("typed-cache fixture");
    let link = LinkSpec::new(typed_source("EmptyCache", "Data.G1"), CachedTable::new())
        .expect("empty cache is a valid authored table");

    let mut edit = source.edit();
    edit.add_link(link).expect("stage empty-cache link");
    let committed = edit
        .commit(&context("ods-dde-empty-cache"))
        .expect("commit empty-cache link");
    let cache = committed.snapshot().links()[2].cached_table_xml();
    assert!(cache.contains("<table:table-row><table:table-cell/></table:table-row>"));
    assert!(!cache.contains("<text:p>"));
    assert!(dde::Snapshot::parse(committed.snapshot().source_xml()).is_ok());
}

#[test]
fn authored_attribute_controls_use_numeric_xml_references() {
    let content = typed_cache_content();
    let source = dde::Snapshot::parse(&content).expect("typed-cache fixture");
    let application = "never\tcontacted\nDDE\rapp";
    let topic = "file:///never\tcontacted\nDDE\rtopic.ods";
    let name = "Captured\tname\nwith\rcontrols";
    let text = "cached\ttext\nwith\rcontrols";
    let source_descriptor = dde::Source::new(application, topic, "Data.A1")
        .expect("control whitespace is valid XML character data")
        .named(name)
        .expect("control whitespace is valid in source names")
        .with_conversion_mode(ConversionMode::KeepText)
        .with_automatic_update(AutomaticUpdate::Disabled);
    let cell = CachedCell::new(CachedValue::Text(text.to_owned())).expect("cache text");
    let row = CachedRow::from_cells(vec![cell]).expect("cache row");
    let cache = CachedTable::from_rows(vec![row]).expect("cache table");
    let link = LinkSpec::new(source_descriptor, cache).expect("complete typed link");

    let mut edit = source.edit();
    edit.add_link(link).expect("stage control-whitespace link");
    let committed = edit
        .commit(&context("ods-dde-control-whitespace"))
        .expect("commit control-whitespace link");
    let xml = committed.snapshot().source_xml();

    assert!(xml.contains("office:dde-application=\"never&#9;contacted&#10;DDE&#13;app\""));
    assert!(xml.contains("office:dde-topic=\"file:///never&#9;contacted&#10;DDE&#13;topic.ods\""));
    assert!(xml.contains("office:name=\"Captured&#9;name&#10;with&#13;controls\""));
    assert!(xml.contains("office:string-value=\"cached&#9;text&#10;with&#13;controls\""));

    let reparsed = dde::Snapshot::parse(xml).expect("numeric character references parse");
    let reparsed_link = &reparsed.links()[2];
    let reparsed_source = reparsed_link.source();
    assert_eq!(reparsed_source.application(), application);
    assert_eq!(reparsed_source.topic(), topic);
    assert_eq!(reparsed_source.name(), Some(name));
    assert!(
        reparsed_link
            .cached_table_xml()
            .contains("cached&#9;text&#10;with&#13;controls")
    );
}

#[test]
fn replacing_a_freshly_authored_link_with_the_same_spec_is_an_exact_noop() {
    let mut content = typed_cache_content();
    let start = content.find("<table:dde-links").expect("link owner start");
    let end = content[start..]
        .find("</table:dde-links>")
        .map(|offset| start + offset + "</table:dde-links>".len())
        .expect("link owner end");
    content.replace_range(start..end, "");

    let source = dde::Snapshot::parse(&content).expect("DDE fixture without links");
    assert!(source.links().is_empty());
    let spec = typed_link("Fresh", "Data.F1");
    let mut add = source.edit();
    add.add_link(spec.clone()).expect("stage fresh link");
    let authored = add
        .commit(&context("ods-dde-fresh-link-authoring"))
        .expect("commit fresh link");
    let authored_source = authored.snapshot().source_xml().to_owned();

    let mut replace = authored.snapshot().edit();
    replace
        .replace_link_at(0usize, spec)
        .expect("replace fresh link with identical spec");
    let no_op = replace
        .commit(&context("ods-dde-fresh-link-identical"))
        .expect("identical fresh-link replacement");
    assert!(!no_op.changed());
    assert!(no_op.patch().is_empty());
    assert_eq!(no_op.snapshot().source_xml(), authored_source);
}

#[test]
fn source_only_move_and_add_in_one_edit_preserves_the_moved_cache() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("composable DDE fixture");
    let first_cache = source.links()[0].cached_table_xml().to_owned();
    let second_cache = source.links()[1].cached_table_xml().to_owned();

    let mut edit = source.edit();
    edit.replace_link_source(0usize, typed_source("MovedSource", "Data.M9"))
        .expect("stage source-only change");
    edit.move_link(0usize, 1usize)
        .expect("move changed existing link");
    edit.add_link(typed_link("AddedAfterMove", "Data.N9"))
        .expect("add link after move");
    let committed = edit
        .commit(&context("ods-dde-source-move-add"))
        .expect("compose source-only move and add");

    let links = committed.snapshot().links();
    assert_eq!(links.len(), 3);
    assert_eq!(links[0].source().item(), "Second.C3");
    assert_eq!(links[0].cached_table_xml(), second_cache);
    assert_eq!(links[1].source().name(), Some("MovedSource"));
    assert_eq!(links[1].source().item(), "Data.M9");
    assert_eq!(links[1].cached_table_xml(), first_cache);
    assert_eq!(links[2].source().name(), Some("AddedAfterMove"));
}

#[test]
fn reordering_cannot_bypass_opaque_cache_replacement_refusal() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("opaque-cache fixture");
    let before = source.source_xml().to_owned();
    let mut edit = source.edit();
    edit.move_link(0usize, 1usize)
        .expect("move opaque link behind clear link");
    edit.replace_link_at(1usize, typed_link("ShouldRefuse", "Data.Z9"))
        .expect("stage replacement at moved position");
    let error = edit
        .commit(&context("ods-dde-opaque-after-reorder"))
        .expect_err("originating opaque cache must still refuse replacement");
    assert!(error.to_string().contains("unknown markup"));
    assert_eq!(edit.before().source_xml(), before);
}

#[test]
fn opaque_cache_replacement_refuses_before_publication() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("opaque-cache fixture");
    let before = source.source_xml().to_owned();
    let mut edit = source.edit();
    edit.replace_link_at(0usize, typed_link("Changed", "Data.Z9"))
        .expect("typed replacement is staged");
    let error = edit
        .commit(&context("ods-dde-opaque-cache"))
        .expect_err("opaque cache replacement must be refused");
    assert!(error.to_string().contains("unknown markup"));
    assert_eq!(edit.before().source_xml(), before);
}

#[test]
fn opaque_cache_text_entities_are_retained_and_refuse_typed_replacement() {
    let content = clean_link_owner_content().replace(
        "<table:table-cell office:value-type=\"float\" office:value=\"7\" table:number-columns-repeated=\"2\"/>",
        "<table:table-cell office:value-type=\"string\"><text:p>A&amp;B&#10;C</text:p></table:table-cell>",
    );
    let source = dde::Snapshot::parse(&content).expect("entity-bearing opaque cache fixture");
    let cache_before = source.links()[0].cached_table_xml().to_owned();
    assert!(cache_before.contains("A&amp;B&#10;C"));

    let mut source_only = source.edit();
    source_only
        .replace_link_source(0usize, typed_source("EntitySource", "Data.E2"))
        .expect("stage source-only entity fixture edit");
    let source_commit = source_only
        .commit(&context("ods-dde-entity-source-only"))
        .expect("source-only edit retains entity-bearing cache");
    assert_eq!(
        source_commit.snapshot().links()[0].cached_table_xml(),
        cache_before
    );

    let before = source_commit.snapshot().source_xml().to_owned();
    let mut replacement = source_commit.snapshot().edit();
    replacement
        .replace_link_at(0usize, typed_link("ShouldRefuseEntities", "Data.E3"))
        .expect("stage typed entity-cache replacement");
    let error = replacement
        .commit(&context("ods-dde-entity-cache-replacement"))
        .expect_err("typed replacement must refuse opaque cache text");
    assert!(error.to_string().contains("unknown markup"));
    assert_eq!(replacement.before().source_xml(), before);
}

#[test]
fn source_only_link_edit_preserves_opaque_cache_bytes() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("opaque-cache fixture");
    let cache_before = source.links()[0].cached_table_xml().to_owned();
    let mut edit = source.edit();
    edit.replace_link_source(0usize, typed_source("SourceOnly", "Data.Z9"))
        .expect("stage source-only replacement");
    let commit = edit
        .commit(&context("ods-dde-source-only"))
        .expect("source-only replacement preserves cache");
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().links()[0].source().name(),
        Some("SourceOnly")
    );
    assert_eq!(commit.snapshot().links()[0].source().item(), "Data.Z9");
    assert_eq!(
        commit.snapshot().links()[0].cached_table_xml(),
        cache_before
    );
    assert!(
        commit.snapshot().links()[0]
            .cached_table_xml()
            .contains("opaque cache comment")
    );
    assert!(
        commit.snapshot().links()[0]
            .cached_table_xml()
            .contains("cache-opaque keep")
    );
}

#[test]
fn repeated_source_only_replacement_preserves_staged_source_and_inverse() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse(&content).expect("opaque-cache fixture");
    let before = source.source_xml().to_owned();
    let source_a = typed_source("SourceA", "Data.A1");
    let source_b = typed_source("SourceB", "Data.B1");
    let mut edit = source.edit();
    edit.replace_link_source(0usize, source_a)
        .expect("stage first source-only replacement");
    edit.replace_link_source(0usize, source_b.clone())
        .expect("stage second source-only replacement");
    edit.replace_link_source(0usize, source_b)
        .expect("repeat the same staged source-only replacement");
    assert!(!edit.is_noop());

    let committed = edit
        .commit(&context("ods-dde-repeated-source-only"))
        .expect("commit the final staged source-only replacement");
    assert_eq!(
        committed.snapshot().links()[0].source().name(),
        Some("SourceB")
    );
    assert_eq!(committed.snapshot().links()[0].source().item(), "Data.B1");
    let restored = committed
        .patch()
        .inverse()
        .apply(committed.snapshot())
        .expect("inverse repeated source-only replacement");
    assert_eq!(restored.snapshot().source_xml(), before);
}

#[test]
fn failed_commit_can_be_retried_without_losing_staged_edit() {
    let content = clean_link_owner_content();
    let source_budget = managed_context("ods-dde-retry-source");
    let other_context = context("ods-dde-retry-other");
    let source =
        dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &source_budget.2)
            .expect("source context snapshot");
    let mut edit = source.edit();
    edit.set_sheet_source("Data", typed_source("Retry", "Data.R1"))
        .expect("stage retry source");

    assert!(edit.commit(&other_context).is_err());
    assert!(!edit.is_noop(), "failed admission retains staged changes");
    let retry = edit
        .commit(&source_budget.2)
        .expect("retry with retained source context");
    assert!(retry.changed());
}

#[test]
fn cancellation_is_checked_before_staging_and_commit_is_atomic() {
    let content = clean_link_owner_content();
    let (_budget, cancellation, context) = managed_context("ods-dde-cancel");
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &context)
        .expect("context snapshot");
    let mut untouched = source.edit();
    cancellation.cancel();
    assert!(
        untouched
            .set_sheet_source("Data", typed_source("Cancelled", "Data.X1"))
            .is_err()
    );
    assert!(untouched.is_noop());

    let (_budget, cancellation, context) = managed_context("ods-dde-cancel-commit");
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &context)
        .expect("context snapshot");
    let mut edit = source.edit();
    edit.set_sheet_source("Data", typed_source("Cancelled", "Data.X1"))
        .expect("stage before cancellation");
    let before = source.source_xml().to_owned();
    cancellation.cancel();
    assert!(edit.commit(&context).is_err());
    assert!(!edit.is_noop());
    assert_eq!(source.source_xml(), before);
}

#[test]
fn output_and_link_limits_fail_without_mutating_the_source() {
    let content = clean_link_owner_content();
    let source = dde::Snapshot::parse_with(
        &content,
        dde::Limits::default()
            .with_links(2)
            .with_output_bytes(content.len()),
    )
    .expect("bounded source snapshot");
    let mut add = source.edit();
    assert!(add.add_link(typed_link("OverLimit", "Data.L1")).is_err());
    assert!(add.is_noop());

    let mut replace = source.edit();
    replace
        .set_sheet_source("Data", typed_source(&"long-name-".repeat(2_000), "Data.L1"))
        .expect("stage oversized source");
    let error = replace
        .commit(&context("ods-dde-output-limit"))
        .expect_err("output limit");
    assert!(error.to_string().contains("OutputBytes") || error.to_string().contains("output"));
    assert_eq!(source.source_xml(), content);
}

#[test]
fn detached_patch_retains_budget_memory_until_final_drop() {
    let content = clean_link_owner_content();
    let (budget, _cancellation, context) = managed_context("ods-dde-patch-retention");
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &context)
        .expect("managed DDE snapshot");
    let after_scan = budget.used(Resource::Memory);
    let mut edit = source.edit();
    edit.set_sheet_source("Data", typed_source("Retained", "Data.R1"))
        .expect("stage retained edit");
    let commit = edit.commit(&context).expect("managed DDE commit");
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    drop(commit);
    drop(edit);
    drop(source);

    assert!(
        budget.used(Resource::Memory) > after_scan,
        "detached patch must retain source and target allocations"
    );
    drop(inverse);
    assert!(
        budget.used(Resource::Memory) > 0,
        "forward patch must retain its source and target allocations"
    );
    drop(patch);
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "final detached patch drop releases retained allocations"
    );
}

#[test]
fn large_typed_cache_admission_is_bounded_and_atomic() {
    let content = clean_link_owner_content();
    let (budget, _cancellation, execution) =
        bounded_context("ods-dde-large-cache-admission", 512 * 1024);
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &execution)
        .expect("reading a small DDE snapshot should fit the finite memory profile");
    let baseline_memory = budget.used(Resource::Memory);
    let before = source.source_xml().to_owned();
    let original_link_count = source.links().len();
    let mut edit = source.edit();
    let error = match edit.add_link(large_typed_link(
        "LargeCache",
        "Data.XFD1",
        65_536,
        "large-cache-cell-payload",
    )) {
        Err(error) => error,
        Ok(_) => panic!("a large typed cache must be refused by the finite memory profile"),
    };
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("memory"),
        "large-cache refusal should identify memory admission: {error}"
    );
    assert_eq!(edit.link_count(), original_link_count);
    assert!(
        edit.is_noop(),
        "failed cache admission must leave the edit unchanged"
    );
    assert_eq!(edit.before().source_xml(), before);
    assert_eq!(
        budget.used(Resource::Memory),
        baseline_memory,
        "a refused payload must release every temporary reservation"
    );
}

#[test]
fn staged_cache_payload_is_retained_until_remove_and_replacements_release_old_payload() {
    let content = clean_link_owner_content();
    let (budget, _cancellation, execution) =
        bounded_context("ods-dde-cache-payload-lifetime", 16 * 1024 * 1024);
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &execution)
        .expect("DDE snapshot should fit the finite memory profile");
    let baseline_memory = budget.used(Resource::Memory);
    let mut edit = source.edit();

    edit.add_link(large_typed_link(
        "StageA",
        "Data.XFD1",
        32_768,
        "cache-payload-cell",
    ))
    .expect("sufficient finite memory should admit the first cache payload");
    let after_add = budget.used(Resource::Memory);
    let retained_payload = after_add.saturating_sub(baseline_memory);
    assert!(
        retained_payload > 128 * 1024,
        "the fixture must exercise a material retained cache payload"
    );

    edit.replace_link_at(
        2usize,
        large_typed_link("StageB", "Data.XFD1", 32_768, "cache-payload-cell"),
    )
    .expect("replacing a staged cache should release the previous payload reservation");
    let after_first_replace = budget.used(Resource::Memory);
    assert!(
        after_first_replace <= after_add + retained_payload / 2,
        "replacement must not retain a second full cache payload"
    );

    edit.replace_link_at(
        2usize,
        large_typed_link("StageC", "Data.XFD1", 32_768, "cache-payload-cell"),
    )
    .expect("repeated replacement should remain within the finite profile");
    let after_second_replace = budget.used(Resource::Memory);
    assert!(
        after_second_replace <= after_add + retained_payload / 2,
        "repeated replacement must release each superseded payload"
    );

    edit.remove_link_at(2usize)
        .expect("removing the staged link should release its cache payload");
    let after_remove = budget.used(Resource::Memory);
    assert!(
        after_remove < after_add,
        "removing the link must release the retained cache allocation"
    );
    drop(edit);
    assert_eq!(
        budget.used(Resource::Memory),
        baseline_memory,
        "dropping the edit must release its remaining draft reservations"
    );
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_only_long_payload_admission_is_atomic_and_bounded() {
    let content = clean_link_owner_content();
    let (tight_budget, _tight_cancellation, tight_execution) =
        bounded_context("ods-dde-source-only-tight", 64 * 1024);
    let tight_source =
        dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &tight_execution)
            .expect("the small source snapshot should fit the tight profile");
    let tight_baseline = tight_budget.used(Resource::Memory);
    let tight_before = tight_source.source_xml().to_owned();
    let mut tight_edit = tight_source.edit();
    let error = match tight_edit.replace_link_source(0usize, long_source_only_source()) {
        Err(error) => error,
        Ok(_) => panic!("the long source-only payload must exceed the tight memory profile"),
    };
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("memory"),
        "source-only refusal should identify memory admission: {error}"
    );
    assert!(tight_edit.is_noop());
    assert_eq!(tight_edit.before().source_xml(), tight_before);
    assert_eq!(tight_budget.used(Resource::Memory), tight_baseline);

    let (sufficient_budget, _sufficient_cancellation, sufficient_execution) =
        bounded_context("ods-dde-source-only-sufficient", 512 * 1024);
    let sufficient_source =
        dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &sufficient_execution)
            .expect("the source snapshot should fit the sufficient profile");
    let sufficient_baseline = sufficient_budget.used(Resource::Memory);
    let mut sufficient_edit = sufficient_source.edit();
    sufficient_edit
        .replace_link_source(0usize, long_source_only_source())
        .expect("sufficient finite memory should admit the source-only payload");
    assert!(!sufficient_edit.is_noop());
    assert!(sufficient_budget.used(Resource::Memory) > sufficient_baseline);
    drop(sufficient_edit);
    assert_eq!(
        sufficient_budget.used(Resource::Memory),
        sufficient_baseline
    );
    drop(sufficient_source);
    assert_eq!(sufficient_budget.used(Resource::Memory), 0);
}

#[test]
fn many_link_scan_commits_and_inverts_under_a_finite_memory_profile() {
    let content = many_links_content(1_024);
    let (budget, _cancellation, execution) =
        bounded_context("ods-dde-many-link-scan-success", 64 * 1024 * 1024);
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &execution)
        .expect("many-link snapshot should fit the finite memory profile");
    assert_eq!(source.links().len(), 1_024);

    let mut edit = source.edit();
    edit.move_link(0usize, 1_023usize)
        .expect("stage a many-link reorder");
    let committed = edit
        .commit(&execution)
        .expect("many-link scan and reorder should fit the finite profile");
    assert!(committed.changed());
    assert_eq!(committed.snapshot().links().len(), 1_024);
    assert_eq!(
        committed.snapshot().links()[1_023].source().name(),
        Some("Link0")
    );
    assert_eq!(
        committed.snapshot().links()[0].source().name(),
        Some("Link1")
    );

    let restored = committed
        .patch()
        .inverse()
        .apply(committed.snapshot())
        .expect("many-link inverse should apply exactly");
    assert_eq!(restored.snapshot().source_xml(), content);
    drop(restored);
    drop(committed);
    drop(edit);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn many_link_scan_memory_refusal_releases_temporary_admission_atomically() {
    let content = many_links_content(1_024);
    let (budget, _cancellation, execution) =
        bounded_context("ods-dde-many-link-scan-tight", 2 * 1024 * 1024);
    let source = dde::Snapshot::parse_with_context(&content, dde::Limits::default(), &execution)
        .expect("many-link snapshot should fit before commit admission");
    assert_eq!(source.links().len(), 1_024);
    let mut edit = source.edit();
    edit.move_link(0usize, 1_023usize)
        .expect("many-link reorder should fit before scan admission");
    let staged_memory = budget.used(Resource::Memory);
    let before = source.source_xml().to_owned();
    let error = match edit.commit(&execution) {
        Err(error) => error,
        Ok(_) => panic!("the tight profile must refuse many-link scan scratch memory"),
    };
    let message = error.to_string().to_ascii_lowercase();
    assert!(
        message.contains("memory"),
        "many-link refusal should identify memory admission: {error}"
    );
    assert!(
        !edit.is_noop(),
        "the staged reorder must remain available for retry"
    );
    assert_eq!(edit.before().source_xml(), before);
    assert_eq!(budget.used(Resource::Memory), staged_memory);
    drop(edit);
    drop(source);
    assert_eq!(budget.used(Resource::Memory), 0);
}
