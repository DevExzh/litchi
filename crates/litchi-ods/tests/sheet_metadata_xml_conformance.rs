//! XML grammar and namespace regressions for the sheet metadata owner.

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits, Profile,
};
use litchi_ods::model::label_range::Orientation;
use litchi_ods::sheet_metadata::{
    CellRange, Detective, LabelRange, Operation, OperationKind, Options, Snapshot,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";

fn context() -> (CancellationSource, ExecutionContext) {
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("nonzero worker count"),
        NonZeroUsize::new(1).expect("nonzero task count"),
        NonZeroU64::new(4096).expect("nonzero byte count"),
        0,
    )
    .expect("finite execution limits");
    let budget = Budget::root(
        "sheet-metadata-xml-conformance",
        BudgetLimits::for_profile(Profile::Server),
    );
    (cancellation, ExecutionContext::new(budget, token, limits))
}

fn options() -> Options {
    Options::new("sum", vec!["Data.A1:Data.A1".to_owned()], "Data.B1").expect("valid consolidation")
}

fn label() -> LabelRange {
    LabelRange::new("Data.A1:Data.A2", "Data.B1:Data.B2", Orientation::Row)
        .expect("valid label range")
}

fn source() -> CellRange {
    CellRange::new("Import", "source.ods#Data.A1", 1, 1).expect("valid source")
}

fn detective() -> Detective {
    let mut value = Detective::new();
    value.add_operation(Operation::new(OperationKind::TraceDependents, 1));
    value
}

fn paired_empty_content() -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}"><office:body><office:spreadsheet><table:label-ranges><table:label-range table:label-cell-range-address="Data.A1:Data.A2" table:data-cell-range-address="Data.B1:Data.B2" table:orientation="row"></table:label-range></table:label-ranges><table:table table:name="Data"><table:table-row><table:table-cell><table:cell-range-source table:name="Import" table:last-column-spanned="1" table:last-row-spanned="1" xlink:type="simple" xlink:href="source.ods#Data.A1"></table:cell-range-source></table:table-cell></table:table-row></table:table><table:consolidation table:function="sum" table:source-cell-range-addresses="Data.A1:Data.A1" table:target-cell-address="Data.B1"></table:consolidation></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[test]
fn paired_empty_owners_are_read_and_noop_exact() {
    let (_cancellation, execution) = context();
    let content = paired_empty_content();
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("paired empty owners are valid XML empty productions");
    assert!(
        snapshot
            .consolidation()
            .expect("consolidation read")
            .is_some()
    );
    assert!(snapshot.label_ranges().is_present());
    assert_eq!(snapshot.label_ranges().len(), 1);
    assert!(
        snapshot
            .cell_metadata(litchi_ods::sheet_metadata::CellSelector::by_name(
                "Data", 0, 0
            ))
            .expect("cell read")
            .expect("cell")
            .range_source()
            .is_some()
    );

    let mut edit = snapshot.edit();
    let commit = edit.commit(&execution).expect("exact no-op");
    assert!(!commit.changed());
    assert_eq!(commit.snapshot().source_xml(), content);

    let mut source_rewrite = snapshot.edit();
    let selector = litchi_ods::sheet_metadata::CellSelector::by_name("Data", 0, 0);
    let mut changed_source = source();
    changed_source.set_name("ChangedImport");
    source_rewrite
        .set_cell_range_source(selector, changed_source)
        .expect("stage paired source rewrite");
    assert!(source_rewrite.commit(&execution).is_err());

    let mut label_rewrite = snapshot.edit();
    label_rewrite
        .add_label_range(label())
        .expect("stage paired label rewrite");
    assert!(label_rewrite.commit(&execution).is_err());
}

#[test]
fn paired_empty_comments_are_retained_for_noop_and_refused_for_rewrite() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}"><office:body><office:spreadsheet><table:table table:name="Data"><table:table-row><table:table-cell><table:cell-range-source table:name="Import" table:last-column-spanned="1" table:last-row-spanned="1" xlink:type="simple" xlink:href="source.ods#Data.A1"><!--keep--></table:cell-range-source></table:table-cell></table:table-row></table:table><table:consolidation table:function="sum" table:source-cell-range-addresses="Data.A1:Data.A1" table:target-cell-address="Data.B1"><!--keep--><?keep yes?></table:consolidation></office:spreadsheet></office:body></office:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("comments and PIs are ignored by empty productions");
    let mut noop = snapshot.edit();
    assert_eq!(
        noop.commit(&execution)
            .expect("no-op")
            .snapshot()
            .source_xml(),
        content
    );

    let mut rewrite = snapshot.edit();
    let mut changed = options();
    changed.function = "average".to_owned();
    rewrite
        .set_consolidation(Some(changed))
        .expect("stage rewrite");
    assert!(rewrite.commit(&execution).is_err());
    assert_eq!(snapshot.source_xml(), content);
}

#[test]
fn root_office_children_follow_schema_order_and_unique_body() {
    let (_cancellation, execution) = context();
    let valid = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body><office:spreadsheet/></office:body></office:document-content>"#
    );
    Snapshot::parse_with_context(
        &valid,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("legal office siblings");

    let duplicate = valid.replace(
        "<office:body><office:spreadsheet/></office:body>",
        "<office:body><office:spreadsheet/></office:body><office:body><office:spreadsheet/></office:body>",
    );
    assert!(
        Snapshot::parse_with_context(
            &duplicate,
            litchi_ods::sheet_metadata::Limits::default(),
            &execution,
        )
        .is_err()
    );

    let out_of_order = valid.replace(
        "<office:automatic-styles/><office:body>",
        "<office:body><office:spreadsheet/></office:body><office:automatic-styles/><office:body>",
    );
    assert!(
        Snapshot::parse_with_context(
            &out_of_order,
            litchi_ods::sheet_metadata::Limits::default(),
            &execution,
        )
        .is_err()
    );
}

#[test]
fn labels_precede_epilogue_owners_without_a_table() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}"><office:body><office:spreadsheet><table:named-expressions/><table:database-ranges/><table:data-pilot-tables/></office:spreadsheet></office:body></office:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("epilogue-only spreadsheet");
    let mut edit = snapshot.edit();
    edit.add_label_range(label()).expect("stage label");
    let commit = edit
        .commit(&execution)
        .expect("insert labels before epilogue");
    let output = commit.snapshot().source_xml();
    assert!(
        output.find("label-ranges").expect("labels")
            < output
                .find("<table:named-expressions/>")
                .expect("named expressions")
    );
    assert!(
        output
            .find("<table:database-ranges/>")
            .expect("database ranges")
            < output
                .find("<table:data-pilot-tables/>")
                .expect("data pilots")
    );
}

#[test]
fn inserted_owners_remain_namespace_valid_at_an_aliased_spreadsheet_context() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:table t:name="Data"><t:table-row><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("aliased document");
    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options()))
        .expect("stage consolidation");
    let commit = edit.commit(&execution).expect("insert aliased owner");
    let output = commit.snapshot().source_xml();
    assert!(output.contains("consolidation"));
    Snapshot::parse_with_context(
        output,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("aliased output reopens");
}

#[test]
fn default_table_elements_get_named_attribute_binding() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet xmlns="{TABLE}"><t:table t:name="Data"><t:table-row><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("default table namespace document");
    let mut edit = snapshot.edit();
    edit.ensure_label_ranges().expect("stage labels");
    let commit = edit
        .commit(&execution)
        .expect("insert default element owner");
    let output = commit.snapshot().source_xml();
    assert!(output.contains("label-ranges"));
    Snapshot::parse_with_context(
        output,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("default namespace output reopens");
}

#[test]
fn local_xlink_alias_is_used_and_nested_declarations_do_not_satisfy_root() {
    let (_cancellation, execution) = context();
    let aliased = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{XLINK}"><o:body><o:spreadsheet><t:table t:name="Data"><t:table-row><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &aliased,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("local alias document");
    let selector = litchi_ods::sheet_metadata::CellSelector::by_name("Data", 0, 0);
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(selector, source())
        .expect("stage aliased source");
    let commit = edit.commit(&execution).expect("insert aliased source");
    let output = commit.snapshot().source_xml();
    assert!(output.contains("cell-range-source"));
    assert!(output.contains("type=\"simple\""));

    let nested = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:s="{TEXT}"><o:body><o:spreadsheet><t:table t:name="Data"><t:table-row><t:table-cell><s:p xmlns:xlink="{XLINK}">text</s:p></t:table-cell></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let nested_snapshot = Snapshot::parse_with_context(
        &nested,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("nested alias document");
    let mut nested_edit = nested_snapshot.edit();
    nested_edit
        .set_cell_range_source(selector, source())
        .expect("stage nested declaration source");
    let nested_commit = nested_edit
        .commit(&execution)
        .expect("insert source and bind xlink at cell");
    let nested_output = nested_commit.snapshot().source_xml();
    let owner_start = nested_output
        .find("cell-range-source")
        .map(|offset| nested_output[..offset].rfind('<').expect("owner opening"))
        .expect("source owner");
    let owner_end = nested_output[owner_start..]
        .find('>')
        .map(|offset| owner_start + offset)
        .expect("owner opening end");
    let owner_opening = &nested_output[owner_start..=owner_end];
    assert!(owner_opening.contains("xmlns:xlink=\"http://www.w3.org/1999/xlink\""));
}

#[test]
fn child_local_table_alias_is_preserved_or_rebound_on_rewrite() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}"><o:body><o:spreadsheet><t:label-ranges><q:label-range xmlns:q="{TABLE}" q:label-cell-range-address="Data.A1:Data.A2" q:data-cell-range-address="Data.B1:Data.B2" q:orientation="row"/></t:label-ranges><t:table t:name="Data"><t:table-row><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("child-local table alias document");
    assert_eq!(snapshot.label_ranges().len(), 1);
    let mut noop = snapshot.edit();
    assert_eq!(
        noop.commit(&execution)
            .expect("child-local no-op")
            .snapshot()
            .source_xml(),
        content
    );

    // The owner is intentionally non-canonical because its child uses a
    // declaration local to that child. A changed operation must preserve the
    // exact source or report an explicit lexical-preservation refusal.
    let mut rewrite = snapshot.edit();
    rewrite
        .add_label_range(label())
        .expect("stage child-local rewrite");
    assert!(rewrite.commit(&execution).is_err());
}

#[test]
fn shadowed_table_alias_does_not_make_inserted_metadata_unbound() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:s="urn:example:foreign"><o:body><o:spreadsheet><t:table t:name="Data"><t:table-row><t:table-cell><s:wrapper xmlns:t="urn:example:foreign"><s:item/></s:wrapper></t:table-cell></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("shadowed alias document");
    let mut edit = snapshot.edit();
    edit.ensure_label_ranges().expect("stage labels");
    let commit = edit.commit(&execution).expect("insert labels");
    Snapshot::parse_with_context(
        commit.snapshot().source_xml(),
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("shadowed alias output remains namespace-valid");
}

#[test]
fn long_table_prefix_does_not_overrun_small_scratch_profile() {
    let (_cancellation, execution) = context();
    let prefix = format!("p{}", "x".repeat(4096));
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:{prefix}="{TABLE}"><o:body><o:spreadsheet><{prefix}:table {prefix}:name="Data"><{prefix}:table-row><{prefix}:table-cell/></{prefix}:table-row></{prefix}:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let limits =
        litchi_ods::sheet_metadata::Limits::new(128 * 1024, 128 * 1024, 32, 1024, 1_000_000)
            .and_then(|limits| limits.with_output_limits(128 * 1024, 1024))
            .expect("finite small scratch profile");
    let snapshot = Snapshot::parse_with_context(&content, limits, &execution)
        .expect("long-prefix input fits profile");
    let mut edit = snapshot.edit();
    edit.ensure_label_ranges().expect("stage labels");
    let commit = edit
        .commit(&execution)
        .expect("new owners use compact canonical prefixes");
    assert!(commit.snapshot().source_xml().len() <= limits.max_output_bytes());
    assert!(commit.snapshot().source_xml().len() - content.len() <= 1024);
    Snapshot::parse_with_context(commit.snapshot().source_xml(), limits, &execution)
        .expect("compact long-prefix output reopens");
}

#[test]
fn locally_declared_known_owners_remain_editable_after_reopen() {
    let (_cancellation, execution) = context();
    let content = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TABLE}" xmlns:x="{XLINK}"><o:body><o:spreadsheet><t:table t:name="Data"><t:table-row><t:table-cell/></t:table-row></t:table></o:spreadsheet></o:body></o:document-content>"#
    );
    let snapshot = Snapshot::parse_with_context(
        &content,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("base document");
    let selector = litchi_ods::sheet_metadata::CellSelector::by_name("Data", 0, 0);
    let mut first = snapshot.edit();
    first
        .set_consolidation(Some(options()))
        .expect("stage consolidation");
    first.ensure_label_ranges().expect("stage labels");
    first
        .set_cell_range_source(selector, source())
        .expect("stage source");
    first
        .set_detective(selector, detective())
        .expect("stage detective");
    let first_commit = first.commit(&execution).expect("first commit");

    let mut second = first_commit.snapshot().edit();
    let mut changed = options();
    changed.function = "average".to_owned();
    second
        .set_consolidation(Some(changed))
        .expect("rewrite consolidation");
    second.add_label_range(label()).expect("rewrite label");
    let mut changed_source = source();
    changed_source.set_name("ImportAgain");
    second
        .set_cell_range_source(selector, changed_source)
        .expect("rewrite source");
    second
        .edit_detective(selector, |value| {
            value.add_operation(Operation::new(OperationKind::TraceErrors, 2));
            Ok(())
        })
        .expect("rewrite detective");
    let second_commit = second.commit(&execution).expect("owners remain editable");
    let output = second_commit.snapshot().source_xml();
    assert!(output.contains("average"));
    assert!(output.contains("ImportAgain"));
    Snapshot::parse_with_context(
        output,
        litchi_ods::sheet_metadata::Limits::default(),
        &execution,
    )
    .expect("second output reopens");
}
