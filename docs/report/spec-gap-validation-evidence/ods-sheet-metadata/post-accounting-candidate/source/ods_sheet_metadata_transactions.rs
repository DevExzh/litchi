//! Public lifecycle evidence for the bounded ODS sheet-metadata owner.
//!
//! The fixtures in this target are intentionally small synthetic ODF 1.4
//! documents.  The local corpus survey did not contain these four metadata
//! families, so every test keeps the exact source XML in the test and labels
//! package members that are unrelated to `content.xml`.

#![allow(
    dead_code,
    reason = "fixture variants are consumed across lifecycle cases"
)]

mod support;

use std::{
    io,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
};

use std::num::{NonZeroU64, NonZeroUsize};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as BudgetLimits,
    OwnedSource, Position, Profile, ReadAt, Resource, SourceVersion,
};
use litchi_odf_common::{
    core::{
        SourceContentPublicationError, SourceContentPublicationOptions,
        SourceContentPublicationProgress,
    },
    package::raw_identical_members,
};
use litchi_ods::model::{
    consolidation::{self, Options, UseLabels},
    detective::{Detective, Direction, HighlightedRange, Operation, OperationKind},
    label_range::{self, Orientation, Range},
    source::CellRange,
};
use litchi_ods::{
    Builder, MutableSpreadsheet, SourceBackedSpreadsheet, Spreadsheet, sheet_metadata as metadata,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TABLE: &str = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK: &str = "http://www.w3.org/1999/xlink";
const VENDOR: &str = "urn:example:vendor";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

const FULL_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<?producer before?>
<office:document-content
    xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:xlink="http://www.w3.org/1999/xlink"
    xmlns:vendor="urn:example:vendor"
    office:version="1.4">
  <!-- document comment retained by focused edits -->
  <office:body>
    <office:spreadsheet>
      <?spreadsheet before?>
      <table:calculation-settings/>
      <table:content-validations/>
      <!-- the label owner belongs to the prelude -->
      <table:label-ranges>
        <!-- label comment -->
        <table:label-range table:label-cell-range-address="Data.A1:A2" table:data-cell-range-address="Data.B1:B2" table:orientation="row"/>
        <?label keep?>
      </table:label-ranges>
      <table:table table:name="Data">
        <table:table-column table:number-columns-repeated="5"/>
        <table:table-row>
          <table:table-cell office:value-type="string">
            <table:cell-range-source table:name="Import" table:last-column-spanned="2" table:last-row-spanned="3" xlink:type="simple" xlink:href="source.ods#Data.A1:B3" xlink:actuate="onRequest" table:filter-name="csv"/>
            <office:annotation><text:p>note</text:p></office:annotation>
            <table:detective>
              <table:highlighted-range table:cell-range-address="Data.B1:B2" table:direction="from-same-table" table:contains-error="false"/>
              <table:highlighted-range table:marked-invalid="true"/>
              <table:operation table:name="trace-errors" table:index="17"/>
              <table:operation table:name="trace-precedents" table:index="2"/>
            </table:detective>
            <text:p>anchor</text:p>
          </table:table-cell>
          <table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>repeat</text:p></table:table-cell>
          <table:covered-table-cell><table:detective/></table:covered-table-cell>
        </table:table-row>
        <table:table-row table:number-rows-repeated="3">
          <table:table-cell table:number-columns-repeated="4" office:value-type="string"><text:p>row-repeat</text:p></table:table-cell>
          <table:covered-table-cell table:number-columns-repeated="2"/>
        </table:table-row>
        <table:table-row>
          <table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2" office:value-type="string"><text:p>merge-anchor</text:p></table:table-cell>
          <table:covered-table-cell/><table:covered-table-cell/>
        </table:table-row>
        <table:table-row><table:covered-table-cell/><table:covered-table-cell/></table:table-row>
      </table:table>
      <table:table table:name="Second"><table:table-row><table:table-cell><text:p>second</text:p></table:table-cell></table:table-row></table:table>
      <?spreadsheet after tables?>
      <table:named-expressions/>
      <table:database-ranges/>
      <table:data-pilot-tables/>
      <table:consolidation table:function="sum" table:source-cell-range-addresses="Data.A1:A2 Data.B1:B2" table:target-cell-address="Data.C1" table:use-labels="both" table:link-to-source-data="true"/>
      <table:dde-links/>
      <vendor:foreign-sibling vendor:keep="yes"><vendor:value>opaque</vendor:value></vendor:foreign-sibling>
      <?producer after?>
    </office:spreadsheet>
  </office:body>
</office:document-content>
"#;

const EMPTY_LABELS_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:spreadsheet><table:label-ranges/><!-- before table --><table:table table:name="Data"><table:table-row><table:table-cell><text:p>empty</text:p></table:table-cell></table:table-row></table:table><!-- after table --></office:spreadsheet></office:body></office:document-content>
"#;

const ABSENT_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4"><office:body><office:spreadsheet><!-- no metadata owners --><table:table table:name="Data"><table:table-row><table:table-cell><text:p>absent</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>
"#;

const ALIASED_CONTENT: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:text:1.0" o:version="1.4"><o:body><o:spreadsheet><t:label-ranges><t:label-range t:label-cell-range-address="Data.A1:A2" t:data-cell-range-address="Data.B1:B2" t:orientation="column"/></t:label-ranges><t:table t:name="Data"><t:table-row><t:table-cell><t:cell-range-source t:name="Alias" t:last-column-spanned="1" t:last-row-spanned="1" x:type="simple" x:href="alias.ods#Data.A1"/></t:table-cell></t:table-row></t:table><t:consolidation t:function="custom" t:source-cell-range-addresses="Data.A1:A2" t:target-cell-address="Data.B1"/></o:spreadsheet></o:body></o:document-content>
"#;

#[allow(dead_code, reason = "fixture variants are consumed by lifecycle cases")]
fn package(content: &str, signed: bool) -> Vec<u8> {
    let mut entries = vec![
        ("content.xml", content.as_bytes(), "text/xml"),
        (
            "Pictures/unrelated.bin",
            b"unrelated-member-bytes".as_slice(),
            "application/octet-stream",
        ),
        (
            "Thumbnails/thumbnail.txt",
            b"thumbnail-comment".as_slice(),
            "text/plain",
        ),
    ];
    if signed {
        entries.push((
            "META-INF/documentsignatures.xml",
            br#"<ds:document-signatures xmlns:ds="urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0"><sig:Signature xmlns:sig="http://www.w3.org/2000/09/xmldsig#"><sig:SignedInfo><sig:CanonicalizationMethod Algorithm="c"/><sig:SignatureMethod Algorithm="s"/><sig:Reference URI="content.xml"><sig:DigestMethod Algorithm="d"/><sig:DigestValue>AAAA</sig:DigestValue></sig:Reference></sig:SignedInfo><sig:SignatureValue>AAAA</sig:SignatureValue></sig:Signature></ds:document-signatures>"#.as_slice(),
            "application/vnd.oasis.opendocument.digital-signature",
        ));
    }
    support::raw_package(&entries)
}

#[allow(dead_code, reason = "fixture variants are consumed by lifecycle cases")]
fn spreadsheet(content: &str) -> Spreadsheet {
    Spreadsheet::from_bytes(package(content, false)).expect("valid synthetic ODS fixture")
}

#[allow(dead_code, reason = "fixture variants are consumed by lifecycle cases")]
fn source_backed(content: &str) -> SourceBackedSpreadsheet {
    SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package(content, false))))
        .expect("valid synthetic source-backed ODS fixture")
}

#[allow(dead_code, reason = "fixture variants are consumed by lifecycle cases")]
fn builder_spreadsheet(content: &str) -> Spreadsheet {
    Spreadsheet::from_bytes(
        Builder::new()
            .content_xml(content)
            .build()
            .expect("valid Builder package"),
    )
    .expect("valid Builder ODS fixture")
}

#[allow(
    dead_code,
    reason = "fixture assertions are shared by ordered-owner cases"
)]
fn assert_in_order(source: &str, needles: &[&str]) {
    let mut cursor = 0;
    for needle in needles {
        let offset = source[cursor..]
            .find(needle)
            .unwrap_or_else(|| panic!("missing ordered XML token {needle:?}"));
        cursor += offset + needle.len();
    }
}

#[allow(dead_code, reason = "fixture rewrites are shared by refusal cases")]
fn replace_once(source: &str, old: &str, new: &str) -> String {
    let mut result = source.to_owned();
    let count = result.matches(old).count();
    assert_eq!(
        count, 1,
        "fixture replacement must identify one source span"
    );
    result = result.replacen(old, new, 1);
    result
}

#[allow(
    dead_code,
    reason = "fixture variants are consumed across grammar cases"
)]
fn repeated_run_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Repeated"><table:table-row table:number-rows-repeated="4"><table:table-cell table:number-columns-repeated="5" office:value-type="string"><text:p>same</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(dead_code, reason = "canonical source is consumed by mutation cases")]
fn canonical_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}" office:version="1.4"><office:body><office:spreadsheet><table:label-ranges><table:label-range table:label-cell-range-address="Data.A1:A2" table:data-cell-range-address="Data.B1:B2" table:orientation="row"/></table:label-ranges><table:table table:name="Data"><table:table-row><table:table-cell office:value-type="string"><table:cell-range-source table:name="Import" table:last-column-spanned="2" table:last-row-spanned="3" xlink:type="simple" xlink:href="source.ods#Data.A1:B3" xlink:actuate="onRequest" table:filter-name="csv"/><table:detective><table:highlighted-range table:cell-range-address="Data.B1:B2" table:direction="from-same-table" table:contains-error="false"/><table:highlighted-range table:marked-invalid="true"/><table:operation table:name="trace-errors" table:index="17"/><table:operation table:name="trace-precedents" table:index="2"/></table:detective><text:p>anchor</text:p></table:table-cell><table:table-cell table:number-columns-repeated="3" office:value-type="string"><text:p>repeat</text:p></table:table-cell></table:table-row></table:table><table:table table:name="Second"><table:table-row><table:table-cell><text:p>second</text:p></table:table-cell></table:table-row></table:table><table:consolidation table:function="sum" table:source-cell-range-addresses="Data.A1:A2 Data.B1:B2" table:target-cell-address="Data.C1" table:use-labels="both" table:link-to-source-data="true"/></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(dead_code, reason = "merge geometry is consumed by selector cases")]
fn implicit_merge_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Merged"><table:table-row><table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="2" office:value-type="string"><text:p>anchor</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(dead_code, reason = "covered geometry is consumed by lifecycle cases")]
fn covered_cell_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Covered"><table:table-row><table:covered-table-cell><table:detective></table:detective></table:covered-table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "fixture variants are consumed by opaque-owner cases"
)]
fn foreign_wrapper_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:vendor="{VENDOR}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Opaque"><table:table-row><table:table-cell><vendor:wrapper><table:detective/></vendor:wrapper><text:p>opaque</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "fixture variants are consumed by namespace/MCE cases"
)]
fn mce_wrapper_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:mc="{MC}" xmlns:vendor="{VENDOR}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Mce"><table:table-row><table:table-cell><mc:AlternateContent><mc:Choice Requires="vendor"><table:detective/></mc:Choice><mc:Fallback/></mc:AlternateContent><text:p>branch</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "row-container grammar is consumed by ordinal cases"
)]
fn grouped_row_containers_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="Grouped"><table:table-header-rows><table:table-row><table:table-cell><text:p>header</text:p></table:table-cell></table:table-row></table:table-header-rows><table:table-row-group><table:table-row><table:table-cell><text:p>group</text:p></table:table-cell></table:table-row></table:table-row-group><table:table-rows><table:table-row><table:table-cell><text:p>rows</text:p></table:table-cell></table:table-row></table:table-rows></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "root grammar refusals are consumed by scanner cases"
)]
fn foreign_wrapped_spreadsheet_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:vendor="{VENDOR}" office:version="1.4"><office:body><vendor:wrapper><office:spreadsheet><table:table table:name="Wrapped"><table:table-row><table:table-cell><text:p>wrapped</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></vendor:wrapper></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "root grammar refusals are consumed by scanner cases"
)]
fn duplicate_spreadsheet_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="First"><table:table-row><table:table-cell><text:p>first</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet><office:spreadsheet><table:table table:name="Second"><table:table-row><table:table-cell><text:p>second</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "root grammar refusals are consumed by scanner cases"
)]
fn direct_non_whitespace_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet>unexpected text<table:table table:name="Data"><table:table-row><table:table-cell><text:p>value</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "hidden-owner refusals are consumed by scanner cases"
)]
fn hidden_detective_under_text_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="HiddenText"><table:table-row><table:table-cell><text:p>visible<table:detective/></text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[allow(
    dead_code,
    reason = "hidden-owner refusals are consumed by scanner cases"
)]
fn hidden_source_under_annotation_content() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="{OFFICE}" xmlns:table="{TABLE}" xmlns:text="{TEXT}" xmlns:xlink="{XLINK}" office:version="1.4"><office:body><office:spreadsheet><table:table table:name="HiddenAnnotation"><table:table-row><table:table-cell><office:annotation><text:p>note<table:cell-range-source table:name="Hidden" table:last-column-spanned="1" table:last-row-spanned="1" xlink:type="simple" xlink:href="hidden.ods#Data.A1"/></text:p></office:annotation><text:p>visible</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document-content>"#
    )
}

#[test]
fn metadata_snapshot_reads_four_families_and_name_position_selectors() {
    let snapshot = metadata::Snapshot::parse(FULL_CONTENT).expect("metadata scan");
    let consolidation = snapshot
        .consolidation()
        .expect("consolidation read")
        .expect("consolidation present");
    assert_eq!(consolidation.function, "sum");
    assert_eq!(consolidation.source_cell_range_addresses.len(), 2);
    assert_eq!(consolidation.use_labels, Some(UseLabels::Both));

    let labels = snapshot.label_ranges();
    assert!(labels.is_present());
    assert_eq!(labels.len(), 1);
    assert_eq!(labels.as_slice()[0].orientation, Orientation::Row);

    let by_name = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Data", 0, 0))
        .expect("name selector")
        .expect("physical anchor");
    assert_eq!(by_name.range_source().expect("source").name(), "Import");
    assert_eq!(by_name.range_source().expect("source").rows(), 3);
    assert!(by_name.range_source().expect("source").actuate_on_request());
    assert_eq!(
        by_name.range_source().expect("source").filter_name(),
        Some("csv")
    );
    assert_eq!(
        by_name.detective().expect("detective").operations()[0].index,
        17
    );
    let location = by_name.location().expect("physical location");
    assert_eq!(location.row(), 0);
    assert_eq!(location.column(), 0);
    assert_eq!(location.row_repeat(), 1);
    assert_eq!(location.column_repeat(), 1);

    let by_position = snapshot
        .cell_metadata(metadata::CellSelector::by_position(Position::new(1), 0, 0))
        .expect("position selector")
        .expect("second sheet cell");
    assert_eq!(by_position.location().expect("location").row(), 0);
    assert!(by_position.range_source().is_none());
    assert!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_name("Missing", 0, 0))
            .is_err()
    );
}

#[test]
fn metadata_snapshot_distinguishes_absent_empty_and_explicit_covered_states() {
    let absent = metadata::Snapshot::parse(ABSENT_CONTENT).expect("absent metadata scan");
    assert!(
        absent
            .consolidation()
            .expect("absent consolidation")
            .is_none()
    );
    assert!(!absent.label_ranges().is_present());
    assert!(absent.label_ranges().is_empty());

    let empty = metadata::Snapshot::parse(EMPTY_LABELS_CONTENT).expect("empty metadata scan");
    assert!(
        empty
            .consolidation()
            .expect("empty consolidation")
            .is_none()
    );
    assert!(empty.label_ranges().is_present());
    assert!(empty.label_ranges().is_empty());

    let covered = metadata::Snapshot::parse(FULL_CONTENT).expect("covered metadata scan");
    let view = covered
        .cell_metadata(metadata::CellSelector::by_name("Data", 0, 4))
        .expect("covered selector")
        .expect("explicit covered cell");
    assert!(view.location().expect("covered location").merge() == litchi_ods::Merge::Covered);
    assert!(view.detective().expect("empty detective owner").is_empty());
}

#[test]
fn metadata_snapshot_reports_implicit_merge_coordinates_without_projecting_anchor_metadata() {
    let snapshot = metadata::Snapshot::parse(&implicit_merge_content()).expect("merge scan");
    let anchor = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Merged", 0, 0))
        .expect("anchor selector")
        .expect("anchor");
    assert_eq!(
        anchor.location().expect("anchor location").merge(),
        litchi_ods::Merge::Span {
            rows: NonZeroUsize::new(2).expect("rows"),
            columns: NonZeroUsize::new(2).expect("columns"),
        }
    );
    assert!(matches!(
        snapshot
            .cell_metadata(metadata::CellSelector::by_name("Merged", 1, 1))
            .expect("implicit covered selector")
            .expect("implicit covered view"),
        metadata::CellMetadataView::ImplicitCovered {
            anchor: (0, 0),
            span: (2, 2)
        }
    ));
    let target = metadata::CellSelector::by_name("Merged", 1, 1);
    let mut edit = snapshot.edit();
    assert!(
        edit.set_cell_range_source(target, range_source("implicit"))
            .is_err()
    );
    assert!(edit.is_no_op());
}

#[test]
fn metadata_edit_can_replace_and_clear_detective_on_an_explicit_covered_cell() {
    let source = metadata::Snapshot::parse(&covered_cell_content()).expect("covered scan");
    let selector = metadata::CellSelector::by_name("Covered", 0, 0);
    let mut edit = source.edit();
    edit.set_detective(selector, detective_value())
        .expect("stage covered detective");
    let commit = edit.commit(&context()).expect("replace covered detective");
    assert!(commit.changed());
    assert_eq!(
        commit
            .snapshot()
            .cell_metadata(selector)
            .expect("covered read")
            .expect("covered physical cell")
            .detective()
            .expect("covered detective")
            .operations()[0]
            .index,
        23
    );

    let mut clear = commit.snapshot().edit();
    clear
        .clear_detective(selector)
        .expect("stage covered detective clear");
    let cleared = clear.commit(&context()).expect("clear covered detective");
    assert!(cleared.changed());
    assert!(
        cleared
            .snapshot()
            .cell_metadata(selector)
            .expect("covered reread")
            .expect("covered physical cell")
            .detective()
            .is_none()
    );
}

#[test]
fn metadata_edit_crud_preserves_order_and_returns_exact_inverse_patch() {
    let source = canonical_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let selector = metadata::CellSelector::by_name("Data", 0, 0);
    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options("average")))
        .expect("replace consolidation");
    edit.insert_label_range(0, label(Orientation::Column))
        .expect("insert first label");
    edit.replace_label_range(1, label(Orientation::Row))
        .expect("replace second label");
    edit.set_cell_range_source(selector, range_source("EditedImport"))
        .expect("replace source");
    edit.edit_detective(selector, |value| {
        value.add_operation(Operation::new(OperationKind::TraceErrors, 99));
        Ok(())
    })
    .expect("edit detective");

    let commit = edit.commit(&context()).expect("metadata commit");
    assert!(commit.changed());
    assert!(
        commit
            .snapshot()
            .source_xml()
            .contains("table:function=\"average\"")
    );
    assert_in_order(
        commit.snapshot().source_xml(),
        &[
            "<table:label-ranges>",
            "<table:label-range",
            "</table:label-ranges>",
            "<table:table",
            "<table:consolidation",
        ],
    );
    let target = commit
        .snapshot()
        .cell_metadata(selector)
        .expect("target cell read")
        .expect("target cell");
    assert_eq!(
        target.range_source().expect("target source").name(),
        "EditedImport"
    );
    assert_eq!(
        target
            .detective()
            .expect("target detective")
            .operations()
            .last()
            .expect("new operation")
            .index,
        99
    );
    assert_eq!(commit.snapshot().label_ranges().len(), 2);

    let applied = commit.patch().apply(&snapshot).expect("patch apply");
    assert_eq!(
        applied.snapshot().source_xml(),
        commit.snapshot().source_xml()
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())
        .expect("inverse apply");
    assert_eq!(restored.snapshot().source_xml(), source);
    let unrelated = metadata::Snapshot::parse(ABSENT_CONTENT).expect("unrelated snapshot");
    assert!(commit.patch().apply(&unrelated).is_err());
}

#[test]
fn metadata_edit_reverts_source_and_detective_to_exact_noop() {
    let absent = metadata::Snapshot::parse(ABSENT_CONTENT).expect("absent metadata scan");
    let absent_selector = metadata::CellSelector::by_name("Data", 0, 0);
    let mut absent_edit = absent.edit();
    absent_edit
        .set_cell_range_source(absent_selector, range_source("Transient"))
        .expect("stage transient source");
    absent_edit
        .clear_cell_range_source(absent_selector)
        .expect("clear transient source");
    absent_edit
        .set_detective(absent_selector, detective_value())
        .expect("stage transient detective");
    absent_edit
        .clear_detective(absent_selector)
        .expect("clear transient detective");
    assert!(absent_edit.is_no_op());
    let absent_commit = absent_edit
        .commit(&context())
        .expect("reverted absent edit commit");
    assert!(!absent_commit.changed());
    assert!(absent_commit.patch().is_empty());
    assert_eq!(absent_commit.snapshot().source_xml(), ABSENT_CONTENT);

    let source = canonical_content();
    let base = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let base_selector = metadata::CellSelector::by_name("Data", 0, 0);
    let base_view = base
        .cell_metadata(base_selector)
        .expect("base cell read")
        .expect("base cell");
    let base_source = base_view.range_source().expect("base source").clone();
    let base_detective = base_view.detective().expect("base detective").clone();
    let mut base_edit = base.edit();
    base_edit
        .set_cell_range_source(base_selector, range_source("Transient"))
        .expect("stage replacement source");
    base_edit
        .set_cell_range_source(base_selector, base_source)
        .expect("restore base source");
    base_edit
        .set_detective(base_selector, detective_value())
        .expect("stage replacement detective");
    base_edit
        .set_detective(base_selector, base_detective)
        .expect("restore base detective");
    assert!(base_edit.is_no_op());
    let base_commit = base_edit
        .commit(&context())
        .expect("reverted base edit commit");
    assert!(!base_commit.changed());
    assert!(base_commit.patch().is_empty());
    assert_eq!(base_commit.snapshot().source_xml(), source);
}

#[test]
fn metadata_edit_detective_callbacks_accumulate_staged_values() {
    let source = canonical_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let selector = metadata::CellSelector::by_name("Data", 0, 0);
    let base_count = snapshot
        .cell_metadata(selector)
        .expect("base detective read")
        .expect("base cell")
        .detective()
        .expect("base detective")
        .operations()
        .len();
    let mut edit = snapshot.edit();
    edit.edit_detective(selector, |value| {
        value.add_operation(Operation::new(OperationKind::TraceErrors, 101));
        Ok(())
    })
    .expect("first detective callback");
    edit.edit_detective(selector, |value| {
        value.add_operation(Operation::new(OperationKind::TraceDependents, 202));
        Ok(())
    })
    .expect("second detective callback");

    let commit = edit
        .commit(&context())
        .expect("accumulated detective commit");
    let operations = commit
        .snapshot()
        .cell_metadata(selector)
        .expect("committed detective read")
        .expect("committed cell")
        .detective()
        .expect("committed detective")
        .operations();
    assert_eq!(operations.len(), base_count + 2);
    assert_eq!(operations[operations.len() - 2].index, 101);
    assert_eq!(operations[operations.len() - 1].index, 202);
}

#[test]
fn metadata_edit_canonicalizes_name_and_position_selectors_to_one_cell() {
    let source = canonical_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let by_name = metadata::CellSelector::by_name("Data", 0, 0);
    let by_position = metadata::CellSelector::by_position(Position::new(0), 0, 0);
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(by_name, range_source("MixedSelector"))
        .expect("stage source by name");
    edit.edit_detective(by_position, |value| {
        value.add_operation(Operation::new(OperationKind::TraceErrors, 303));
        Ok(())
    })
    .expect("stage detective by position");

    let commit = edit.commit(&context()).expect("mixed selector commit");
    let target = commit
        .snapshot()
        .cell_metadata(by_name)
        .expect("mixed selector readback")
        .expect("mixed selector physical cell");
    assert_eq!(
        target.range_source().expect("mixed source").name(),
        "MixedSelector"
    );
    assert_eq!(
        target
            .detective()
            .expect("mixed detective")
            .operations()
            .last()
            .expect("mixed operation")
            .index,
        303
    );
    assert_eq!(
        commit
            .snapshot()
            .cell_metadata(by_position)
            .expect("position readback")
            .expect("position physical cell")
            .range_source()
            .expect("position source")
            .name(),
        "MixedSelector"
    );
}

#[test]
fn metadata_edit_mixed_selector_revert_returns_exact_noop() {
    let source = canonical_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let by_name = metadata::CellSelector::by_name("Data", 0, 0);
    let by_position = metadata::CellSelector::by_position(Position::new(0), 0, 0);
    let original = snapshot
        .cell_metadata(by_name)
        .expect("original source read")
        .expect("original cell")
        .range_source()
        .expect("original source")
        .clone();
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(by_name, range_source("Transient"))
        .expect("stage replacement by name");
    edit.set_cell_range_source(by_position, original)
        .expect("restore source by position");
    assert!(edit.is_no_op());
    let commit = edit
        .commit(&context())
        .expect("mixed selector no-op commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), source);
}

#[test]
fn metadata_commit_requires_the_snapshot_execution_context_lineage() {
    let source = canonical_content();
    let (_source_budget, _source_cancellation, source_context) =
        managed_context("ods-sheet-metadata-lineage-source");
    let (_other_budget, _other_cancellation, other_context) =
        managed_context("ods-sheet-metadata-lineage-other");
    let snapshot = metadata::Snapshot::parse_with_context(
        &source,
        metadata::Limits::default(),
        &source_context,
    )
    .expect("managed lineage snapshot");
    let selector = metadata::CellSelector::by_name("Data", 0, 0);
    let mut edit = snapshot.edit();
    edit.set_detective(selector, detective_value())
        .expect("stage lineage edit");
    assert!(
        edit.commit(&other_context).is_err(),
        "a commit from a different context lineage must be refused"
    );
    assert!(
        !edit.is_no_op(),
        "failed context commit must retain staging"
    );
    let commit = edit
        .commit(&source_context)
        .expect("same-context lineage commit");
    assert!(commit.changed());
}

#[test]
fn metadata_index_traverses_valid_row_containers_with_logical_ordinals() {
    let source = grouped_row_containers_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("grouped row-container scan");
    for row in 0..3 {
        let selector = metadata::CellSelector::by_name("Grouped", row, 0);
        let view = snapshot
            .cell_metadata(selector)
            .expect("grouped row selector")
            .expect("grouped physical cell");
        let location = view.location().expect("grouped location");
        assert_eq!(location.row(), row);
        assert_eq!(location.column(), 0);
        assert_eq!(location.row_repeat(), 1);
        assert_eq!(location.column_repeat(), 1);
    }

    let mut edit = snapshot.edit();
    edit.set_cell_range_source(
        metadata::CellSelector::by_name("Grouped", 0, 0),
        range_source("Header"),
    )
    .expect("edit header-row cell");
    edit.set_detective(
        metadata::CellSelector::by_name("Grouped", 1, 0),
        detective_value(),
    )
    .expect("edit row-group cell");
    edit.set_cell_range_source(
        metadata::CellSelector::by_name("Grouped", 2, 0),
        range_source("Rows"),
    )
    .expect("edit table-rows cell");
    let commit = edit
        .commit(&context())
        .expect("commit grouped row-container edits");
    assert!(commit.changed());
    assert!(commit.snapshot().source_xml().contains("table-header-rows"));
    assert!(commit.snapshot().source_xml().contains("table-row-group"));
    assert!(commit.snapshot().source_xml().contains("table-rows"));
    assert_eq!(
        commit
            .snapshot()
            .cell_metadata(metadata::CellSelector::by_name("Grouped", 0, 0))
            .expect("header readback")
            .expect("header cell")
            .range_source()
            .expect("header source")
            .name(),
        "Header"
    );
    assert!(
        commit
            .snapshot()
            .cell_metadata(metadata::CellSelector::by_name("Grouped", 1, 0))
            .expect("group readback")
            .expect("group cell")
            .detective()
            .is_some()
    );
    assert_eq!(
        commit
            .snapshot()
            .cell_metadata(metadata::CellSelector::by_name("Grouped", 2, 0))
            .expect("rows readback")
            .expect("rows cell")
            .range_source()
            .expect("rows source")
            .name(),
        "Rows"
    );
}

#[test]
fn metadata_edit_staging_is_failure_atomic_and_rollback_restores_presence() {
    let snapshot = metadata::Snapshot::parse(EMPTY_LABELS_CONTENT).expect("empty labels scan");
    let mut edit = snapshot.edit();
    let before = edit.staged_label_ranges().as_slice().to_vec();
    assert!(
        edit.replace_label_range(1, label(Orientation::Column))
            .is_err()
    );
    assert_eq!(edit.staged_label_ranges().as_slice(), before.as_slice());
    assert!(
        edit.set_cell_range_source(
            metadata::CellSelector::by_name("Missing", 0, 0),
            range_source("missing"),
        )
        .is_err()
    );
    assert!(edit.is_no_op());

    edit.clear_label_ranges().expect("clear empty owner");
    assert!(!edit.staged_label_ranges().is_present());
    edit.rollback();
    assert!(edit.staged_label_ranges().is_present());
    assert!(edit.staged_label_ranges().is_empty());
    let commit = edit.commit(&context()).expect("rolled-back no-op");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), snapshot.source_xml());
}

#[test]
fn metadata_edit_creates_and_clears_absent_singletons_without_collapsing_empty_labels() {
    let snapshot = metadata::Snapshot::parse(ABSENT_CONTENT).expect("absent metadata scan");
    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options("sum")))
        .expect("add consolidation");
    edit.ensure_label_ranges()
        .expect("add empty label container");
    let commit = edit.commit(&context()).expect("add absent owners");
    assert!(
        commit
            .snapshot()
            .consolidation()
            .expect("consolidation")
            .is_some()
    );
    assert!(commit.snapshot().label_ranges().is_present());
    assert!(commit.snapshot().label_ranges().is_empty());
    assert_in_order(
        commit.snapshot().source_xml(),
        &[
            "<table:label-ranges/>",
            "<table:table",
            "<table:consolidation",
        ],
    );

    let mut clear = commit.snapshot().edit();
    clear.clear_consolidation().expect("clear consolidation");
    clear.clear_label_ranges().expect("clear labels");
    let cleared = clear.commit(&context()).expect("clear owners");
    assert!(
        cleared
            .snapshot()
            .consolidation()
            .expect("cleared consolidation")
            .is_none()
    );
    assert!(!cleared.snapshot().label_ranges().is_present());
}

#[test]
fn metadata_scanner_uses_namespace_uris_and_keeps_same_named_foreign_wrappers_opaque() {
    let aliased = metadata::Snapshot::parse(ALIASED_CONTENT).expect("aliased metadata scan");
    assert_eq!(
        aliased
            .consolidation()
            .expect("aliased consolidation")
            .expect("owner")
            .function,
        "custom"
    );
    assert_eq!(
        aliased.label_ranges().as_slice()[0].orientation,
        Orientation::Column
    );
    let aliased_cell = aliased
        .cell_metadata(metadata::CellSelector::by_name("Data", 0, 0))
        .expect("aliased cell")
        .expect("aliased physical cell");
    assert_eq!(
        aliased_cell.range_source().expect("aliased source").name(),
        "Alias"
    );

    for content in [foreign_wrapper_content(), mce_wrapper_content()] {
        let snapshot = metadata::Snapshot::parse(&content).expect("opaque wrapper scan");
        let selector = metadata::CellSelector::by_name(
            if content.contains("Mce") {
                "Mce"
            } else {
                "Opaque"
            },
            0,
            0,
        );
        let view = snapshot
            .cell_metadata(selector)
            .expect("opaque cell read")
            .expect("opaque physical cell");
        assert!(view.detective().is_none());
        let mut edit = snapshot.edit();
        assert!(edit.set_detective(selector, detective_value()).is_err());
        assert!(edit.is_no_op());
        assert_eq!(snapshot.source_xml(), content);
    }
}

#[test]
fn metadata_scanner_refuses_malformed_direct_owner_order_and_duplicate_owners_before_staging() {
    let source_after_annotation = replace_once(
        &canonical_content(),
        "<table:cell-range-source",
        "<office:annotation/><table:cell-range-source",
    );
    let snapshot = metadata::Snapshot::parse(&source_after_annotation)
        .expect("malformed ordering remains inspectable");
    let selector = metadata::CellSelector::by_name("Data", 0, 0);
    let mut edit = snapshot.edit();
    assert!(edit.clear_cell_range_source(selector).is_err());
    assert!(edit.is_no_op());

    // Add a second direct owner to exercise the ambiguous write-set refusal.
    let duplicate = replace_once(
        &canonical_content(),
        r#"/><table:detective>"#,
        r#"/><table:cell-range-source table:name="Second" table:last-column-spanned="1" table:last-row-spanned="1" xlink:type="simple" xlink:href="second.ods"/><table:detective>"#,
    );
    let snapshot = metadata::Snapshot::parse(&duplicate).expect("duplicate owner scan");
    let mut edit = snapshot.edit();
    assert!(edit.clear_cell_range_source(selector).is_err());
    assert!(edit.is_no_op());
}

#[test]
fn metadata_scanner_refuses_invalid_roots_and_hidden_cell_owners_before_insertion() {
    for (name, content) in [
        (
            "foreign-wrapped spreadsheet",
            foreign_wrapped_spreadsheet_content(),
        ),
        (
            "duplicate spreadsheet root",
            duplicate_spreadsheet_content(),
        ),
        ("direct non-whitespace", direct_non_whitespace_content()),
    ] {
        assert!(
            metadata::Snapshot::parse(&content).is_err(),
            "invalid {name} must be refused before an index is published"
        );
    }

    let hidden_text = hidden_detective_under_text_content();
    let snapshot = metadata::Snapshot::parse(&hidden_text)
        .expect("hidden detective remains inspectable as opaque source");
    let selector = metadata::CellSelector::by_name("HiddenText", 0, 0);
    let mut edit = snapshot.edit();
    assert!(
        edit.set_detective(selector, detective_value()).is_err(),
        "nested detective under text must refuse before insertion"
    );
    assert!(edit.is_no_op());
    assert_eq!(snapshot.source_xml(), hidden_text);

    let hidden_annotation = hidden_source_under_annotation_content();
    let snapshot = metadata::Snapshot::parse(&hidden_annotation)
        .expect("hidden source remains inspectable as opaque source");
    let selector = metadata::CellSelector::by_name("HiddenAnnotation", 0, 0);
    let mut edit = snapshot.edit();
    assert!(
        edit.set_cell_range_source(selector, range_source("Replacement"))
            .is_err(),
        "nested source under annotation must refuse before insertion"
    );
    assert!(edit.is_no_op());
    assert_eq!(snapshot.source_xml(), hidden_annotation);
}

#[test]
fn metadata_scanner_rejects_unbound_namespaces_duplicate_attributes_and_bad_values() {
    let unbound = replace_once(
        &canonical_content(),
        "<table:label-ranges>",
        "<bad:label-ranges>",
    );
    assert!(metadata::Snapshot::parse(&unbound).is_err());

    let duplicate_attribute = replace_once(
        &canonical_content(),
        "<table:consolidation ",
        "<table:consolidation table:function=\"sum\" ",
    );
    assert!(metadata::Snapshot::parse(&duplicate_attribute).is_err());

    let zero_dimension = replace_once(
        &canonical_content(),
        "table:last-row-spanned=\"3\"",
        "table:last-row-spanned=\"0\"",
    );
    assert!(metadata::Snapshot::parse(&zero_dimension).is_err());

    let invalid_operation = replace_once(
        &canonical_content(),
        "table:name=\"trace-errors\" table:index=\"17\"",
        "table:name=\"invalid-operation\" table:index=\"17\"",
    );
    let negative_operation = replace_once(
        &canonical_content(),
        "table:name=\"trace-errors\" table:index=\"17\"",
        "table:name=\"trace-errors\" table:index=\"-1\"",
    );
    for malformed in [invalid_operation, negative_operation] {
        let snapshot = metadata::Snapshot::parse(&malformed)
            .expect("malformed detective owner remains exact opaque source");
        let selector = metadata::CellSelector::by_name("Data", 0, 0);
        let mut edit = snapshot.edit();
        edit.set_detective(selector, detective_value())
            .expect("stage replacement before preservation refusal");
        assert!(edit.commit(&context()).is_err());
        edit.rollback();
        assert_eq!(snapshot.source_xml(), malformed);
    }
}

#[test]
fn metadata_edit_splits_repeated_row_and_cell_runs_for_one_logical_target() {
    let source = repeated_run_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("repeated metadata scan");
    let target = metadata::CellSelector::by_name("Repeated", 2, 2);
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(target, range_source("Interior"))
        .expect("stage interior source");
    edit.set_detective(target, detective_value())
        .expect("stage interior detective");
    let commit = edit.commit(&context()).expect("split repeated runs");
    let xml = commit.snapshot().source_xml();
    assert!(xml.contains("table:number-rows-repeated=\"2\""));
    assert!(xml.contains("table:number-rows-repeated=\"1\""));
    assert!(xml.contains("table:number-columns-repeated=\"2\""));
    let prefix = commit
        .snapshot()
        .cell_metadata(metadata::CellSelector::by_name("Repeated", 2, 0))
        .expect("prefix read")
        .expect("prefix cell");
    assert!(prefix.range_source().is_none());
    assert!(prefix.detective().is_none());
    let selected = commit
        .snapshot()
        .cell_metadata(target)
        .expect("target read")
        .expect("target cell");
    assert_eq!(
        selected.range_source().expect("target source").name(),
        "Interior"
    );
    assert!(selected.detective().is_some());
    let suffix = commit
        .snapshot()
        .cell_metadata(metadata::CellSelector::by_name("Repeated", 2, 4))
        .expect("suffix read")
        .expect("suffix cell");
    assert!(suffix.range_source().is_none());
    assert!(suffix.detective().is_none());
}

#[test]
fn metadata_edit_merges_two_targets_in_one_repeated_row_and_cell_run() {
    let source = repeated_run_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("repeated metadata scan");
    let source_target = metadata::CellSelector::by_name("Repeated", 2, 1);
    let detective_target = metadata::CellSelector::by_name("Repeated", 2, 3);
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(source_target, range_source("Left"))
        .expect("stage first repeated target");
    edit.set_detective(detective_target, detective_value())
        .expect("stage second repeated target");

    let commit = edit
        .commit(&context())
        .expect("merge repeated-row target replacements");
    assert!(commit.changed());
    let committed = commit.snapshot();
    let xml = committed.source_xml();
    assert!(
        xml.find("table:name=\"Left\"")
            .expect("serialized source owner")
            < xml
                .find("<table:detective>")
                .expect("serialized detective owner")
    );
    assert_eq!(
        committed
            .cell_metadata(source_target)
            .expect("source target read")
            .expect("source target physical cell")
            .range_source()
            .expect("source target owner")
            .name(),
        "Left"
    );
    assert!(
        committed
            .cell_metadata(detective_target)
            .expect("detective target read")
            .expect("detective target physical cell")
            .detective()
            .is_some()
    );

    for column in [0, 2, 4] {
        let view = committed
            .cell_metadata(metadata::CellSelector::by_name("Repeated", 2, column))
            .expect("untouched repeated target read")
            .expect("untouched repeated physical cell");
        assert!(view.range_source().is_none());
        assert!(view.detective().is_none());
    }
    for row in [0, 1, 3] {
        let view = committed
            .cell_metadata(metadata::CellSelector::by_name("Repeated", row, 0))
            .expect("untouched repeated row read")
            .expect("untouched repeated row cell");
        assert!(view.range_source().is_none());
        assert!(view.detective().is_none());
    }

    let restored = commit
        .patch()
        .inverse()
        .apply(committed)
        .expect("inverse two-target repeated patch");
    assert_eq!(restored.snapshot().source_xml(), source);
}

#[test]
fn metadata_noop_on_repeated_runs_retains_exact_source_without_splitting() {
    let source = canonical_content();
    let snapshot = metadata::Snapshot::parse(&source).expect("canonical metadata scan");
    let existing = snapshot
        .cell_metadata(metadata::CellSelector::by_name("Data", 0, 0))
        .expect("existing cell")
        .expect("existing metadata")
        .range_source()
        .expect("existing source")
        .clone();
    let mut edit = snapshot.edit();
    edit.set_cell_range_source(metadata::CellSelector::by_name("Data", 0, 0), existing)
        .expect("stage equal source");
    let commit = edit.commit(&context()).expect("equal metadata commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), source);
}

#[test]
fn managed_patch_retains_source_and_target_memory_until_final_patch_drop() {
    let (budget, _cancellation, context) = managed_context("ods-sheet-metadata-retention");
    let source = canonical_content();
    let snapshot =
        metadata::Snapshot::parse_with_context(&source, metadata::Limits::default(), &context)
            .expect("managed metadata snapshot");
    let after_scan = budget.used(Resource::Memory);
    assert!(
        after_scan >= source.len() as u64,
        "source allocation must be retained under the caller budget"
    );

    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options("average")))
        .expect("stage managed edit");
    let commit = edit.commit(&context).expect("managed metadata commit");
    let after_commit = budget.used(Resource::Memory);
    assert!(
        after_commit > after_scan,
        "accepted target must add retained source memory"
    );

    let patch = commit.patch().clone();
    let patch_clone = patch.clone();
    let inverse = patch.inverse();
    drop(commit);
    drop(edit);
    drop(snapshot);

    let after_origin_drop = budget.used(Resource::Memory);
    assert!(
        after_origin_drop > 0,
        "detached patch values must retain source and target allocations"
    );
    drop(inverse);
    assert!(
        budget.used(Resource::Memory) > 0,
        "cloned forward patch must keep retained memory alive"
    );
    drop(patch_clone);
    assert!(
        budget.used(Resource::Memory) > 0,
        "the final forward patch must own retained memory"
    );
    drop(patch);
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "all retained source and target reservations release with the final patch"
    );
}

#[test]
fn metadata_patch_apply_honors_retained_context_cancellation() {
    let (budget, cancellation, context) = managed_context("ods-sheet-metadata-apply");
    let source = canonical_content();
    let snapshot =
        metadata::Snapshot::parse_with_context(&source, metadata::Limits::default(), &context)
            .expect("managed metadata snapshot");
    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options("average")))
        .expect("stage managed apply edit");
    let commit = edit.commit(&context).expect("managed apply commit");
    let patch = commit.patch().clone();

    cancellation.cancel();
    let error = patch
        .apply(&snapshot)
        .expect_err("cancelled patch apply must stop before default parsing");
    assert!(error.to_string().to_ascii_lowercase().contains("cancel"));
    assert!(
        budget.used(Resource::Memory) > 0,
        "failed cancelled apply must preserve originating retained values"
    );
}

#[test]
fn metadata_cell_staging_checks_selector_cancellation_before_comparing_candidates() {
    let (budget, cancellation, context) = managed_context("ods-sheet-metadata-selector-stage");
    let snapshot = metadata::Snapshot::parse_with_context(
        &canonical_content(),
        metadata::Limits::default(),
        &context,
    )
    .expect("managed metadata snapshot");
    let mut edit = snapshot.edit();
    let before_work = budget.used(Resource::Work);

    cancellation.cancel();
    let error = edit
        .set_cell_range_source(
            metadata::CellSelector::by_name("Data", 0, 0),
            range_source("Data.A1"),
        )
        .expect_err("cancelled selector staging must stop before scanning candidates");
    assert!(error.to_string().to_ascii_lowercase().contains("cancel"));
    assert_eq!(budget.used(Resource::Work), before_work);
    assert!(
        edit.is_no_op(),
        "cancelled staging must leave the edit unchanged"
    );
}

#[test]
fn managed_label_staging_honors_retained_cancellation_and_releases_memory() {
    let (budget, cancellation, context) = managed_context("ods-sheet-metadata-staging");
    let snapshot = metadata::Snapshot::parse_with_context(
        ABSENT_CONTENT,
        metadata::Limits::default(),
        &context,
    )
    .expect("managed absent metadata snapshot");
    let mut edit = snapshot.edit();
    let before_ranges = edit.staged_label_ranges().as_slice().to_vec();
    let before_present = edit.staged_label_ranges().is_present();
    let before_memory = budget.used(Resource::Memory);

    cancellation.cancel();
    let error = edit
        .add_label_range(label(Orientation::Column))
        .expect_err("cancelled staging mutator must stop before allocation");
    assert!(error.to_string().to_ascii_lowercase().contains("cancel"));
    assert_eq!(
        edit.staged_label_ranges().as_slice(),
        before_ranges.as_slice()
    );
    assert_eq!(edit.staged_label_ranges().is_present(), before_present);
    assert!(
        edit.is_no_op(),
        "cancelled staging must leave the public edit ledger unchanged"
    );
    assert_eq!(
        budget.used(Resource::Memory),
        before_memory,
        "cancelled staging must leave the retained budget and ledger unchanged"
    );
    drop(edit);
    drop(snapshot);
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "cancelled staging must release the originating snapshot reservation"
    );

    let (budget, _cancellation, context) = managed_context("ods-sheet-metadata-staging-live");
    let snapshot = metadata::Snapshot::parse_with_context(
        ABSENT_CONTENT,
        metadata::Limits::default(),
        &context,
    )
    .expect("managed live metadata snapshot");
    let mut edit = snapshot.edit();
    let before_staging = budget.used(Resource::Memory);
    edit.add_label_range(label(Orientation::Column))
        .expect("live label staging");
    assert_eq!(edit.staged_label_ranges().len(), 1);
    assert!(
        budget.used(Resource::Memory) > before_staging,
        "accepted staged labels and strings must be retained under the caller budget"
    );
    drop(edit);
    assert_eq!(
        budget.used(Resource::Memory),
        before_staging,
        "staged label reservations must release when the edit is dropped"
    );
    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_backed_metadata_commit_round_trips_inverse_and_untouched_members() {
    let source = package(&canonical_content(), false);
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))
        .expect("source-backed canonical fixture");
    let before = owner.sheet_metadata().expect("source metadata snapshot");
    assert_eq!(
        before
            .consolidation()
            .expect("read source consolidation")
            .expect("source consolidation")
            .function,
        "sum"
    );

    let mut edit = before.edit().expect("source metadata edit");
    edit.set_consolidation(Some(options("average")))
        .expect("stage source consolidation");
    let commit = edit.commit(&context()).expect("source metadata commit");
    assert!(commit.changed());
    assert_eq!(commit.patch().source_xml(), before.source_xml());
    assert_eq!(commit.patch().target_xml(), commit.snapshot().source_xml());

    let applied = commit
        .patch()
        .apply(&before)
        .expect("apply source-backed patch");
    assert_eq!(
        applied.snapshot().source_xml(),
        commit.snapshot().source_xml()
    );
    let applied_via_owner = owner
        .apply_sheet_metadata_patch(commit.patch())
        .expect("apply source-backed patch through facade");
    assert_eq!(
        applied_via_owner.snapshot().source_xml(),
        commit.snapshot().source_xml()
    );
    let restored = commit
        .patch()
        .inverse()
        .apply(applied.snapshot())
        .expect("apply source-backed inverse");
    assert_eq!(restored.snapshot().source_xml(), before.source_xml());

    let mut output = Vec::new();
    let report = commit
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect("publish source-backed metadata commit");
    assert!(!report.is_no_op());
    assert_eq!(report.bytes(), output.len() as u64);
    let identical = raw_identical_members(&source, &output).expect("compare package members");
    assert!(identical.contains("mimetype"));
    assert!(identical.contains("META-INF/manifest.xml"));
    assert!(identical.contains("Pictures/unrelated.bin"));
    assert!(identical.contains("Thumbnails/thumbnail.txt"));
    assert!(!identical.contains("content.xml"));

    let reopened = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(output)))
        .expect("reopen published package");
    assert_eq!(
        reopened
            .sheet_metadata()
            .expect("reopened metadata")
            .consolidation()
            .expect("reopened consolidation")
            .expect("reopened owner")
            .function,
        "average"
    );
}

#[test]
fn source_backed_noop_publication_retains_exact_xml_and_reports_noop() {
    let content = canonical_content();
    let source = package(&content, false);
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(source.clone())))
        .expect("source-backed canonical fixture");
    let before = owner.sheet_metadata().expect("source metadata snapshot");
    let mut edit = before.edit().expect("source metadata edit");
    let commit = edit.commit(&context()).expect("source no-op commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().source_xml(), content);

    let mut output = Vec::new();
    let report = commit
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect("publish source no-op");
    assert!(report.is_no_op());
    assert_eq!(report.bytes(), output.len() as u64);
    let reopened = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(output)))
        .expect("reopen no-op package");
    assert_eq!(reopened.content_xml().expect("reopened content"), content);
    assert!(!source.is_empty());
}

#[test]
fn source_backed_publication_honors_limits_and_cancellation() {
    let owner = source_backed(&canonical_content());
    let mut edit = owner.edit_sheet_metadata().expect("source metadata edit");
    edit.set_consolidation(Some(options("average")))
        .expect("stage source consolidation");
    let commit = edit.commit(&context()).expect("source metadata commit");

    let error = commit
        .write_to(
            Vec::new(),
            SourceContentPublicationOptions::new().with_max_output_bytes(1),
        )
        .expect_err("one-byte output limit");
    assert!(matches!(
        error,
        SourceContentPublicationError::LimitExceeded { .. }
    ));
    assert_eq!(
        error.progress(),
        SourceContentPublicationProgress::Untouched
    );

    let (cancellation, token) = CancellationSource::pair();
    cancellation.cancel();
    let error = commit
        .write_to(
            Vec::new(),
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .expect_err("pre-cancelled publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled {
            progress: SourceContentPublicationProgress::Untouched
        }
    ));

    let limits =
        metadata::Limits::new(1, 1024, 128, 1024, 1_000).expect("small finite metadata limits");
    let limited = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package(
        &canonical_content(),
        false,
    ))))
    .expect("source-backed limited fixture");
    assert!(limited.sheet_metadata_with(limits, &context()).is_err());
}

#[test]
fn source_backed_metadata_detects_stale_source_before_publication() {
    let source = Arc::new(MutableSource::new(package(&canonical_content(), false)));
    let owner = SourceBackedSpreadsheet::from_read_at(source.clone())
        .expect("mutable source-backed fixture");
    let before = owner.sheet_metadata().expect("source metadata snapshot");
    let mut edit = before.edit().expect("source metadata edit");
    edit.set_consolidation(Some(options("average")))
        .expect("stage source consolidation");
    let commit = edit.commit(&context()).expect("source metadata commit");

    source.bump();
    assert!(before.consolidation().is_err());
    assert!(commit.patch().apply(&before).is_err());
    let mut output = Vec::new();
    let error = commit
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect_err("stale source publication");
    assert!(matches!(
        error,
        SourceContentPublicationError::SourceChanged {
            progress: SourceContentPublicationProgress::Untouched,
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn signed_source_refuses_changed_metadata_while_preserving_staged_edit() {
    let owner = SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(package(
        &canonical_content(),
        true,
    ))))
    .expect("signed source-backed fixture");
    let before = owner
        .content_xml()
        .expect("signed source content")
        .to_owned();
    let mut edit = owner
        .edit_sheet_metadata()
        .expect("signed source metadata edit");
    edit.set_consolidation(Some(options("average")))
        .expect("stage signed source edit");
    assert!(edit.commit(&context()).is_err());
    assert!(edit.commit(&context()).is_err());
    assert_eq!(
        owner.content_xml().expect("source remains unchanged"),
        before
    );
}

#[test]
fn ordinary_signed_facades_refuse_changed_metadata_publication() {
    let content = canonical_content();
    let mut signed =
        Spreadsheet::from_bytes(package(&content, true)).expect("signed ordinary fixture");
    let before = signed.content_xml().to_owned();
    let error = signed
        .edit_sheet_metadata(|edit| edit.set_consolidation(Some(options("average"))))
        .expect_err("ordinary signed metadata edit");
    assert!(matches!(error, litchi_core::Error::Unsupported(_)));
    assert_eq!(signed.content_xml(), before);

    let unsigned = spreadsheet(&content);
    let mut edit = unsigned.sheet_metadata().expect("unsigned metadata").edit();
    edit.set_consolidation(Some(options("average")))
        .expect("stage unsigned patch");
    let patch = edit
        .commit(&context())
        .expect("unsigned patch")
        .patch()
        .clone();
    let error = signed
        .apply_sheet_metadata_patch(&patch)
        .expect_err("ordinary signed patch");
    assert!(matches!(error, litchi_core::Error::Unsupported(_)));
    assert_eq!(signed.content_xml(), before);

    signed
        .edit_sheet_metadata(|_| Ok(()))
        .expect("signed no-op remains inspectable");

    let mut mutable =
        MutableSpreadsheet::from_bytes(package(&content, true)).expect("signed mutable fixture");
    let mutable_before = mutable.spreadsheet().content_xml().to_owned();
    let error = mutable
        .edit_sheet_metadata(|edit| edit.set_consolidation(Some(options("average"))))
        .expect_err("mutable signed metadata edit");
    assert!(matches!(error, litchi_core::Error::Unsupported(_)));
    assert_eq!(mutable.spreadsheet().content_xml(), mutable_before);
    mutable
        .edit_sheet_metadata(|_| Ok(()))
        .expect("mutable signed no-op remains inspectable");
}

#[test]
fn ordinary_and_mutable_facades_forward_metadata_snapshot_edit_and_patch() {
    let content = canonical_content();
    let mut ordinary = spreadsheet(&content);
    let before = ordinary
        .sheet_metadata()
        .expect("ordinary metadata snapshot");
    let mut staged = before.edit();
    staged
        .set_consolidation(Some(options("average")))
        .expect("stage ordinary consolidation");
    let commit = staged.commit(&context()).expect("ordinary metadata commit");
    ordinary
        .apply_sheet_metadata_patch(commit.patch())
        .expect("ordinary patch forwarding");
    assert_eq!(
        ordinary
            .sheet_metadata()
            .expect("ordinary reread")
            .consolidation()
            .expect("ordinary consolidation")
            .expect("ordinary owner")
            .function,
        "average"
    );

    ordinary
        .edit_sheet_metadata(|edit| edit.clear_consolidation())
        .expect("ordinary closure forwarding");
    assert!(
        ordinary
            .sheet_metadata()
            .expect("ordinary cleared reread")
            .consolidation()
            .expect("ordinary cleared consolidation")
            .is_none()
    );

    let mut mutable =
        MutableSpreadsheet::from_bytes(package(&content, false)).expect("mutable metadata fixture");
    mutable
        .edit_sheet_metadata(|edit| edit.set_consolidation(Some(options("average"))))
        .expect("mutable closure forwarding");
    assert_eq!(
        mutable
            .sheet_metadata()
            .expect("mutable reread")
            .consolidation()
            .expect("mutable consolidation")
            .expect("mutable owner")
            .function,
        "average"
    );
}

#[test]
fn changed_owner_preserves_comments_processing_instructions_and_foreign_siblings() {
    let mut ordinary = spreadsheet(FULL_CONTENT);
    ordinary
        .edit_sheet_metadata(|edit| edit.set_consolidation(Some(options("average"))))
        .expect("change consolidation in annotated fixture");
    let content = ordinary.content_xml();
    assert!(content.contains("document comment retained"));
    assert!(content.contains("producer before"));
    assert!(content.contains("producer after"));
    assert!(content.contains("foreign-sibling"));
    assert!(content.contains("label comment"));
    assert!(content.contains("<?label keep?>"));
}

struct MutableSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|_| io::Error::other("source too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(bytes) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let length = bytes.len().min(output.len());
        output[..length].copy_from_slice(&bytes[..length]);
        Ok(length)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f44_5353,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[allow(
    dead_code,
    reason = "context helper is used by bounded lifecycle cases"
)]
fn managed_context(scope: &str) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(scope.to_owned(), BudgetLimits::for_profile(Profile::Server));
    let (cancellation, token) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget.clone(),
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).expect("non-zero worker count"),
            NonZeroUsize::new(1).expect("non-zero task count"),
            NonZeroU64::new(1024 * 1024).expect("non-zero byte count"),
            0,
        )
        .expect("valid execution limits"),
    );
    (budget, cancellation, context)
}

#[allow(
    dead_code,
    reason = "context helper is used by bounded lifecycle cases"
)]
fn context() -> ExecutionContext {
    managed_context("ods-sheet-metadata-test").2
}

fn options(function: &str) -> Options {
    Options::new(
        function,
        vec!["Data.A1:A2".to_owned(), "Data.B1:B2".to_owned()],
        "Data.C1",
    )
    .expect("valid consolidation")
}

fn label(orientation: Orientation) -> Range {
    Range::new("Data.C1:C2", "Data.D1:D2", orientation).expect("valid label range")
}

fn range_source(name: &str) -> CellRange {
    let mut value = CellRange::new(name, "source.ods#Data.C1:D3", 3, 2).expect("source");
    value.set_actuate_on_request(true);
    value.set_filter_name(Some("csv".to_owned()));
    value
}

fn detective_value() -> Detective {
    let mut value = Detective::new();
    value
        .add_highlighted_range(
            HighlightedRange::valid(
                Some("Data.C1:D2".to_owned()),
                Direction::FromSameTable,
                Some(true),
            )
            .expect("valid detective range"),
        )
        .add_operation(Operation::new(OperationKind::TraceDependents, 23));
    value
}

#[test]
fn metadata_scanner_accepts_standard_document_content_siblings() {
    let source = canonical_content().replace(
        "<office:body>",
        "<office:scripts/><office:font-face-decls/><office:automatic-styles/><office:body>",
    );
    let snapshot = metadata::Snapshot::parse(&source)
        .expect("standard document-content siblings remain valid metadata input");
    let mut edit = snapshot.edit();
    edit.set_consolidation(Some(options("average")))
        .expect("stage consolidation with root siblings");
    let commit = edit.commit(&context()).expect("commit with root siblings");
    assert!(commit.changed());
    assert!(commit.snapshot().source_xml().contains("<office:scripts/>"));
    assert!(
        commit
            .snapshot()
            .source_xml()
            .contains("<office:font-face-decls/>")
    );
    assert!(
        commit
            .snapshot()
            .source_xml()
            .contains("<office:automatic-styles/>")
    );
}

#[test]
fn existing_public_models_cover_the_four_metadata_values_without_side_effects() {
    let consolidation = consolidation::parse_consolidation(FULL_CONTENT)
        .expect("consolidation owner should parse")
        .expect("fixture has one consolidation");
    assert_eq!(consolidation.function, "sum");
    assert_eq!(consolidation.source_cell_range_addresses.len(), 2);
    assert_eq!(consolidation.use_labels, Some(UseLabels::Both));

    let labels = label_range::parse(FULL_CONTENT).expect("label owner should parse");
    assert_eq!(labels.len(), 1);
    assert_eq!(labels[0].orientation, Orientation::Row);

    let source = CellRange::new("Import", "source.ods#Data.A1:B3", 3, 2)
        .expect("positive source dimensions");
    assert_eq!(source.rows(), 3);
    assert_eq!(source.columns(), 2);

    let mut detective = Detective::new();
    detective
        .add_highlighted_range(
            HighlightedRange::valid(
                Some("Data.B1:B2".to_owned()),
                Direction::FromSameTable,
                Some(false),
            )
            .expect("valid highlighted range"),
        )
        .add_operation(Operation::new(OperationKind::TraceErrors, 17));
    assert_eq!(detective.highlighted_ranges().len(), 1);
    assert_eq!(detective.operations()[0].index, 17);
    let _ = Options::new("sum", vec!["Data.A1:A2".to_owned()], "Data.C1")
        .expect("valid consolidation options");
    let _ = Range::new("Data.A1:A2", "Data.B1:B2", Orientation::Column).expect("valid label range");
}

#[test]
fn aliases_are_resolved_by_expanded_namespace_in_the_existing_public_codecs() {
    let consolidation = consolidation::parse_consolidation(ALIASED_CONTENT)
        .expect("aliased consolidation should parse")
        .expect("aliased fixture has consolidation");
    assert_eq!(consolidation.function, "custom");
    let labels = label_range::parse(ALIASED_CONTENT).expect("aliased labels should parse");
    assert_eq!(labels[0].orientation, Orientation::Column);
    assert!(ALIASED_CONTENT.contains("<t:cell-range-source"));
}

#[test]
fn synthetic_fixture_is_a_complete_ods_package_with_untouched_members() {
    let bytes = package(FULL_CONTENT, false);
    let spreadsheet = Spreadsheet::from_bytes(bytes).expect("open fixture");
    assert_eq!(spreadsheet.sheet_names(), ["Data", "Second"]);
    assert!(spreadsheet.content_xml().contains("producer before"));
    assert!(spreadsheet.content_xml().contains("foreign-sibling"));
}
