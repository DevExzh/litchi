//! Losslessness, dependency, and reversible-edit regression tests.

use std::collections::BTreeMap;

use litchi_sheet::{Cell as Address, Column, Rect, Row};

use super::{
    Action, ColumnAction, DefaultsAction, DescentEffect, HeightEffect, MergePlan, OptionalEffect,
    Plan, RowAction, StyleEffect, WidthEffect, rewrite, rewrite_merges,
};
use crate::cell::{Cell, Value};
use crate::column::Width;
use crate::error::{
    ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, MergeEditBlock, RowEditBlock,
};
use crate::layout::{self, Descent};
use crate::outline::Outline;
use crate::raw::worksheet;
use crate::row::Height;
const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

fn exact_slice<'a>(content: &'a [u8], needle: &[u8]) -> &'a [u8] {
    let start = content
        .windows(needle.len())
        .position(|window| window == needle)
        .expect("expected byte slice");
    &content[start..start + needle.len()]
}

#[test]
fn no_op_and_redundant_merge_plans_preserve_source_bytes() {
    let source = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension z:keep="dimension" ref="A1"/><x:sheetData/><x:mergeCells z:keep="container" count="1"><x:mergeCell z:keep="record" ref="E5:F5"/></x:mergeCells><x:hyperlinks/></x:worksheet>"#
        )
        .into_bytes();

    assert_eq!(
        rewrite(&source, "Data", Plan::default()).expect("no-op edit"),
        source
    );
    assert_eq!(
        rewrite_merges(&source, "Data", MergePlan::default()).expect("no-op merge edit"),
        source
    );

    let existing = Rect::from_a1("E5:F5").expect("existing merge");
    assert_eq!(
        rewrite_merges(
            &source,
            "Data",
            MergePlan {
                add: vec![existing],
                remove: Vec::new(),
            },
        )
        .expect("redundant merge add"),
        source
    );

    let absent = Rect::from_a1("B2:C3").expect("absent merge");
    assert_eq!(
        rewrite_merges(
            &source,
            "Data",
            MergePlan {
                add: Vec::new(),
                remove: vec![absent],
            },
        )
        .expect("redundant merge remove"),
        source
    );
}

#[test]
fn failed_validation_does_not_publish_a_partial_edit() {
    let source = format!(
            r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><sheetProtection sheet="1"/></worksheet>"#
        )
        .into_bytes();
    let original = source.clone();
    let result = rewrite(
        &source,
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::from([(
                Address::from_a1("A1").expect("A1"),
                Action::set(7_i32.into()),
            )]),
            rows: BTreeMap::new(),
            columns: BTreeMap::new(),
        },
    );

    assert!(matches!(
        result,
        Err(Error::EditBlocked {
            reason: EditBlock::ProtectedSheet,
            ..
        })
    ));
    assert_eq!(source, original);
}

#[test]
fn untouched_xml_fragments_remain_byte_exact() {
    let source = format!(
            r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData>
  <x:row r='1' data='edited'><x:c r='A1'><x:v>1</x:v></x:c></x:row>
  <!-- keep this comment and whitespace -->
  <x:row r='2' z:opaque='yes'>
    <x:c r='B2'><x:v>2</x:v></x:c>
  </x:row>
</x:sheetData><x:future z:payload='untouched' /><x:mergeCells count='1'><x:mergeCell z:record='yes' ref='E5:F5'/></x:mergeCells></x:worksheet>"#
        )
        .into_bytes();
    let untouched_row =
        b"<x:row r='2' z:opaque='yes'>\n    <x:c r='B2'><x:v>2</x:v></x:c>\n  </x:row>";
    let untouched_child = b"<x:future z:payload='untouched' />";
    let untouched_merge = b"<x:mergeCell z:record='yes' ref='E5:F5'/>";

    let edited = rewrite(
        &source,
        "Data",
        BTreeMap::from([(
            Address::from_a1("A1").expect("A1"),
            Action::set("updated".into()),
        )]),
    )
    .expect("cell edit");
    assert_eq!(
        exact_slice(&source, untouched_row),
        exact_slice(&edited, untouched_row)
    );
    assert_eq!(
        exact_slice(&source, untouched_child),
        exact_slice(&edited, untouched_child)
    );

    let merge_source = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData>
  <x:row r='1' data='edited'><x:c r='A1'><x:v>1</x:v></x:c></x:row>
  <!-- keep this comment and whitespace -->
  <x:row r='2' z:opaque='yes'>
    <x:c r='B2'><x:v>2</x:v></x:c>
  </x:row>
</x:sheetData><x:mergeCells count='1'><x:mergeCell z:record='yes' ref='E5:F5'/></x:mergeCells></x:worksheet>"#
    )
    .into_bytes();
    let merged = rewrite_merges(
        &merge_source,
        "Data",
        MergePlan {
            add: vec![Rect::from_a1("B2:C3").expect("new merge")],
            remove: Vec::new(),
        },
    )
    .expect("merge edit");
    assert_eq!(
        exact_slice(&merge_source, untouched_row),
        exact_slice(&merged, untouched_row)
    );
    assert_eq!(
        exact_slice(&merge_source, untouched_merge),
        exact_slice(&merged, untouched_merge)
    );
}

#[test]
fn minimally_rewrites_set_clear_remove_and_new_rows() {
    let xml = format!(
        r#"<?xml version="1.0"?><x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension ref="A1:C1" z:hint="kept"/><x:sheetData data="kept">
  <x:row r="1" spans="1:4" z:row="kept"><x:c r="A1" s="2" t="s" z:cell="kept"><x:v>0</x:v><x:extLst><z:data/></x:extLst></x:c><x:c r="C1"><x:v>3</x:v></x:c></x:row>
  <x:row r="5"><x:c r="D5" s="4"/></x:row>
</x:sheetData><x:extLst><z:untouched value="yes"/></x:extLst></x:worksheet>"#
    );
    let mut actions = BTreeMap::new();
    actions.insert(
        Address::from_a1("A1").unwrap(),
        Action::set("new & text".into()),
    );
    actions.insert(Address::from_a1("B1").unwrap(), Action::set(42_i32.into()));
    actions.insert(Address::from_a1("C1").unwrap(), Action::Remove);
    actions.insert(Address::from_a1("D5").unwrap(), Action::clear(true));
    actions.insert(Address::from_a1("A3").unwrap(), Action::set(true.into()));

    let edited = rewrite(xml.as_bytes(), "Data", actions).unwrap();
    let edited = std::str::from_utf8(&edited).unwrap();
    assert!(edited.contains(r#"z:cell="kept""#));
    assert!(edited.contains(r#"<x:dimension z:hint="kept" ref="A1:D5"/>"#));
    assert!(edited.contains("<x:extLst><z:data/></x:extLst>"));
    assert!(edited.contains(
            r#"<x:c s="2" z:cell="kept" r="A1" t="inlineStr"><x:is><x:t xml:space="preserve">new &amp; text</x:t></x:is>"#
        ));
    assert!(edited.contains(r#"<x:c r="B1"><x:v>42</x:v></x:c>"#));
    assert!(!edited.contains(r#"r="C1""#));
    assert!(edited.contains(r#"<x:row r="3"><x:c r="A3" t="b"><x:v>1</x:v></x:c></x:row>"#));
    assert!(edited.contains(r#"<x:c s="4" r="D5"></x:c>"#));
    assert!(edited.contains(r#"<x:extLst><z:untouched value="yes"/></x:extLst>"#));
    assert!(!edited.contains("spans="));

    let store = worksheet::parse(edited.as_bytes(), || Ok(None)).unwrap();
    assert!(matches!(
        store.get(Address::from_a1("A1").unwrap()),
        Some(Cell::Value(Value::Text(text))) if text.as_str() == "new & text"
    ));
    assert!(store.get(Address::from_a1("C1").unwrap()).is_none());
    assert!(matches!(
        store.get(Address::from_a1("D5").unwrap()),
        Some(Cell::Empty)
    ));
}

#[test]
fn dimension_expansion_never_narrows_producer_bounds() {
    let empty = format!(r#"<worksheet xmlns="{S}"><dimension ref="A1"/><sheetData/></worksheet>"#);
    let created = rewrite(
        empty.as_bytes(),
        "Data",
        BTreeMap::from([(
            Address::from_a1("C3").expect("address"),
            Action::set(1_i32.into()),
        )]),
    )
    .expect("create C3");
    assert!(
        std::str::from_utf8(&created)
            .expect("UTF-8")
            .contains(r#"<dimension ref="A1:C3"/>"#)
    );

    let populated = format!(
        r#"<worksheet xmlns="{S}"><dimension ref="A1:C3"/><sheetData><row r="3"><c r="C3"><v>1</v></c></row></sheetData></worksheet>"#
    );
    let removed = rewrite(
        populated.as_bytes(),
        "Data",
        BTreeMap::from([(Address::from_a1("C3").expect("address"), Action::Remove)]),
    )
    .expect("remove C3");
    assert!(
        std::str::from_utf8(&removed)
            .expect("UTF-8")
            .contains(r#"<dimension ref="A1:C3"/>"#)
    );

    let absent = format!(r#"<worksheet xmlns="{S}"><sheetData/></worksheet>"#);
    let edited = rewrite(
        absent.as_bytes(),
        "Data",
        BTreeMap::from([(
            Address::from_a1("B2").expect("address"),
            Action::set(1_i32.into()),
        )]),
    )
    .expect("edit without producer dimension");
    assert!(
        !std::str::from_utf8(&edited)
            .expect("UTF-8")
            .contains("dimension")
    );
}

#[test]
fn row_visibility_surgery_is_sparse_lossless_and_composes_with_cells() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension ref="A1:A4"/><x:sheetData><x:row r="1" hidden="1" z:keep="yes"><x:c r="A1"><x:v>1</x:v></x:c></x:row><x:row r="2" hidden="0" z:empty="keep"/><x:row r="4"><x:c r="A4"><x:v>4</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
    );
    let plan = Plan {
        defaults: None,
        cells: BTreeMap::from([(
            Address::from_a1("A4").expect("A4"),
            Action::set(40_i32.into()),
        )]),
        rows: BTreeMap::from([
            (Row::new(0).expect("row 1"), RowAction::show()),
            (Row::new(1).expect("row 2"), RowAction::hide()),
            (Row::new(2).expect("row 3"), RowAction::hide()),
            (Row::new(3).expect("row 4"), RowAction::hide()),
        ]),
        columns: BTreeMap::new(),
    };

    let edited = rewrite(xml.as_bytes(), "Data", plan).expect("visibility edit");
    let edited = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(edited.contains(r#"<x:row r="1" z:keep="yes">"#));
    assert!(edited.contains(r#"<x:row r="2" z:empty="keep" hidden="1"/>"#));
    assert!(edited.contains(r#"<x:row r="3" hidden="1"/>"#));
    assert!(edited.contains(r#"<x:row r="4" hidden="1"><x:c r="A4"><x:v>40</x:v>"#));
    assert!(edited.contains(r#"<x:dimension ref="A1:A4"/>"#));

    let store = worksheet::parse(edited.as_bytes(), || Ok(None)).expect("reparse rows");
    assert!(!store.row(Row::new(0).expect("row 1")).hidden());
    assert!(store.row(Row::new(1).expect("row 2")).hidden());
    assert!(store.row(Row::new(2).expect("row 3")).hidden());
    assert!(store.row(Row::new(3).expect("row 4")).hidden());
    assert!(matches!(
        store.get(Address::from_a1("A4").expect("A4")),
        Some(Cell::Value(Value::Number(value))) if value.as_str() == "40"
    ));
}

#[test]
fn row_layout_facets_preserve_unedited_state_and_materialize_sparsely() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData><x:row r="2" s="1" customFormat="1" ht="20" customHeight="1" hidden="1" outlineLevel="2" collapsed="1" thickTop="1" thickBot="1" ph="1" z:keep="yes"><x:c r="A2"><x:v>2</x:v></x:c></x:row></x:sheetData></x:worksheet>"#
    );
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::from([
                (
                    Row::new(1).expect("row 2"),
                    RowAction {
                        hidden: Some(false),
                        height: Some(HeightEffect::Reset),
                        outline: Some(Outline::new(3).expect("outline")),
                        collapsed: Some(false),
                        thick_top: Some(false),
                        phonetic: Some(false),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(2).expect("row 3"),
                    RowAction {
                        height: Some(HeightEffect::Set(Height::new(25.0).expect("height"))),
                        outline: Some(Outline::new(1).expect("outline")),
                        collapsed: Some(true),
                        thick_bottom: Some(true),
                        phonetic: Some(true),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(3).expect("row 4"),
                    RowAction {
                        hidden: Some(false),
                        height: Some(HeightEffect::Reset),
                        ..RowAction::default()
                    },
                ),
            ]),
            columns: BTreeMap::new(),
        },
    )
    .expect("row layout rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(concat!(
        r#"<x:row r="2" s="1" customFormat="1" thickBot="1" z:keep="yes" "#,
        r#"outlineLevel="3">"#
    )));
    assert!(text.contains(concat!(
        r#"<x:row r="3" ht="25" customHeight="1" outlineLevel="1" "#,
        r#"collapsed="1" thickBot="1" ph="1"/>"#
    )));
    assert!(!text.contains(r#"r="4""#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse row layout");
    let second = store.row(Row::new(1).expect("row 2"));
    assert_eq!(second.height(), None);
    assert!(!second.custom_height());
    assert!(!second.hidden());
    assert_eq!(second.outline().get(), 3);
    assert!(!second.collapsed());
    assert!(!second.thick_top());
    assert!(second.thick_bottom());
    assert!(!second.phonetic());
    assert!(second.custom_format());
    assert_eq!(
        store.row_entry(second.index()).unwrap().properties.style,
        Some(1)
    );
    let third = store.row(Row::new(2).expect("row 3"));
    assert_eq!(third.height().map(Height::get), Some(25.0));
    assert!(third.custom_height());
    assert_eq!(third.outline().get(), 1);
    assert!(third.collapsed());
    assert!(third.thick_bottom());
    assert!(third.phonetic());
    assert!(!store.row(Row::new(3).expect("row 4")).stored());
}

#[test]
fn worksheet_defaults_and_row_descent_rewrite_losslessly_by_facet() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"
                xmlns:compat="http://schemas.openxmlformats.org/markup-compatibility/2006"
                xmlns:ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"
                compat:Ignorable="ac">
                <x:sheetFormatPr baseColWidth="10" defaultColWidth="12"
                    defaultRowHeight="15" customHeight="0" zeroHeight="1"
                    thickTop="1" ac:dyDescent="0.1" z:keep="yes"/>
                <x:sheetData z:data="keep"><x:row r="1" customHeight="0"
                    ac:dyDescent="0.2" z:row="keep"/><x:row r="2"/></x:sheetData>
            </x:worksheet>"#
    );
    let mut defaults = DefaultsAction::default();
    {
        let effects = defaults.update();
        effects.base_width = Some(OptionalEffect::Reset);
        effects.width = Some(OptionalEffect::Set(
            layout::Width::new(14.5).expect("default width"),
        ));
        effects.height = Some(layout::Height::new(20.0).expect("default height"));
        effects.hidden = Some(false);
        effects.thick_top = Some(false);
        effects.thick_bottom = Some(true);
        effects.descent = Some(DescentEffect::Set(
            Descent::new(0.25).expect("default descent"),
        ));
    }
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: Some(defaults),
            cells: BTreeMap::new(),
            rows: BTreeMap::from([
                (
                    Row::new(0).expect("row 1"),
                    RowAction {
                        descent: Some(DescentEffect::Reset),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(1).expect("row 2"),
                    RowAction {
                        descent: Some(DescentEffect::Set(Descent::new(0.3).expect("row descent"))),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(2).expect("row 3"),
                    RowAction {
                        descent: Some(DescentEffect::Set(Descent::new(0.4).expect("row descent"))),
                        ..RowAction::default()
                    },
                ),
            ]),
            columns: BTreeMap::new(),
        },
    )
    .expect("defaults rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(r#"z:keep="yes""#));
    assert!(text.contains(r#"z:data="keep""#));
    assert!(text.contains(r#"z:row="keep""#));
    assert!(!text.contains("baseColWidth="));
    assert!(!text.contains("zeroHeight="));
    assert!(!text.contains("thickTop="));
    assert!(text.contains(r#"defaultColWidth="14.5""#));
    assert!(text.contains(r#"defaultRowHeight="20" customHeight="1""#));
    assert!(text.contains(r#"thickBottom="1" ac:dyDescent="0.25""#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse defaults");
    let defaults = store.defaults().expect("stored defaults");
    assert_eq!(defaults.stored_base_width(), None);
    assert_eq!(defaults.base_width(), layout::DEFAULT_BASE_WIDTH);
    assert_eq!(defaults.width().map(layout::Width::get), Some(14.5));
    assert_eq!(defaults.height().get(), 20.0);
    assert!(!defaults.hidden());
    assert!(!defaults.thick_top());
    assert!(defaults.thick_bottom());
    assert_eq!(defaults.descent().map(Descent::get), Some(0.25));
    assert_eq!(
        store
            .row(Row::new(0).expect("row 1"))
            .descent()
            .map(Descent::get),
        None
    );
    assert_eq!(
        store
            .row(Row::new(1).expect("row 2"))
            .descent()
            .map(Descent::get),
        Some(0.3)
    );
    assert_eq!(
        store
            .row(Row::new(2).expect("row 3"))
            .descent()
            .map(Descent::get),
        Some(0.4)
    );

    let removed = rewrite(
        &edited,
        "Data",
        Plan {
            defaults: Some(DefaultsAction::remove()),
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::new(),
        },
    )
    .expect("remove defaults");
    assert!(
        worksheet::parse(&removed, || Ok(None))
            .expect("reparse removed defaults")
            .defaults()
            .is_none()
    );
}

#[test]
fn new_descent_injects_collision_free_ignorable_namespaces() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:x14ac="urn:occupied"
                xmlns:mc="urn:also-occupied"><x:sheetData
                xmlns:x14ac1="urn:locally-occupied"><x:row r="5"/></x:sheetData></x:worksheet>"#
    );
    let mut defaults = DefaultsAction::default();
    {
        let effects = defaults.update();
        effects.height = Some(layout::Height::new(17.0).expect("height"));
        effects.descent = Some(DescentEffect::Set(
            Descent::new(0.2).expect("default descent"),
        ));
    }
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: Some(defaults),
            cells: BTreeMap::new(),
            rows: BTreeMap::from([(
                Row::new(4).expect("row 5"),
                RowAction {
                    descent: Some(DescentEffect::Set(Descent::new(0.35).expect("row descent"))),
                    ..RowAction::default()
                },
            )]),
            columns: BTreeMap::new(),
        },
    )
    .expect("inject extension namespaces");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(concat!(
        r#"xmlns:x14ac2="http://schemas.microsoft.com/office/"#,
        r#"spreadsheetml/2009/9/ac""#
    )));
    assert!(text.contains(concat!(
        r#"xmlns:mc1="http://schemas.openxmlformats.org/"#,
        r#"markup-compatibility/2006""#
    )));
    assert!(text.contains(r#"mc1:Ignorable="x14ac2""#));
    assert!(text.contains(r#"x14ac2:dyDescent="0.2""#));
    assert!(text.contains(r#"x14ac2:dyDescent="0.35""#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse injected XML");
    let defaults = store.defaults().expect("materialized defaults");
    assert_eq!(defaults.height().get(), 17.0);
    assert_eq!(defaults.descent().map(Descent::get), Some(0.2));
    assert_eq!(
        store
            .row(Row::new(4).expect("row 5"))
            .descent()
            .map(Descent::get),
        Some(0.35)
    );
}

#[test]
fn defaults_edits_refuse_missing_dependencies_before_rewrite() {
    let plain = format!(r#"<worksheet xmlns="{S}"><sheetData/></worksheet>"#);
    let mut needs_height = DefaultsAction::default();
    needs_height.update().width = Some(OptionalEffect::Set(
        layout::Width::new(12.0).expect("width"),
    ));
    assert!(matches!(
        rewrite(
            plain.as_bytes(),
            "Data",
            Plan {
                defaults: Some(needs_height),
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::new(),
            },
        ),
        Err(Error::DefaultsEditBlocked {
            reason: DefaultsEditBlock::NeedsHeight,
            ..
        })
    ));

    let protected =
        format!(r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#);
    assert!(matches!(
        rewrite(
            protected.as_bytes(),
            "Data",
            Plan {
                defaults: Some(DefaultsAction::remove()),
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::new(),
            },
        ),
        Err(Error::DefaultsEditBlocked {
            reason: DefaultsEditBlock::ProtectedSheet,
            ..
        })
    ));

    let compatibility = format!(
        r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><extLst><ext><sheetFormatPr
                defaultRowHeight="15"/></ext></extLst><sheetData/></worksheet>"#
    );
    assert!(matches!(
        rewrite(
            compatibility.as_bytes(),
            "Data",
            Plan {
                defaults: Some(DefaultsAction::remove()),
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::new(),
            },
        ),
        Err(Error::DefaultsEditBlocked {
            reason: DefaultsEditBlock::MarkupCompatibility,
            ..
        })
    ));
}

#[test]
fn row_style_retargeting_derives_custom_format_and_resets_sparsely() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData><x:row r="2" s="1" customFormat="1" ht="20" z:keep="yes"/></x:sheetData></x:worksheet>"#
    );
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::from([
                (
                    Row::new(1).expect("row 2"),
                    RowAction {
                        style: Some(StyleEffect::Reset),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(2).expect("row 3"),
                    RowAction {
                        style: Some(StyleEffect::Set(2)),
                        ..RowAction::default()
                    },
                ),
                (
                    Row::new(3).expect("row 4"),
                    RowAction {
                        style: Some(StyleEffect::Reset),
                        ..RowAction::default()
                    },
                ),
            ]),
            columns: BTreeMap::new(),
        },
    )
    .expect("row style rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(r#"<x:row r="2" ht="20" z:keep="yes"/>"#));
    assert!(text.contains(r#"<x:row r="3" s="2" customFormat="1"/>"#));
    assert!(!text.contains(r#"r="4""#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse row styles");
    let second = store.row(Row::new(1).expect("row 2"));
    assert_eq!(
        store
            .row_entry(second.index())
            .expect("row 2")
            .properties
            .style,
        None
    );
    assert!(!second.custom_format());
    let third = store.row(Row::new(2).expect("row 3"));
    assert_eq!(
        store
            .row_entry(third.index())
            .expect("row 3")
            .properties
            .style,
        Some(2)
    );
    assert!(third.custom_format());
    assert!(!store.row(Row::new(3).expect("row 4")).stored());
}

#[test]
fn protected_sheet_blocks_row_visibility_before_rewrite() {
    let xml =
        format!(r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#);
    let result = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::from([(Row::new(0).expect("row 1"), RowAction::hide())]),
            columns: BTreeMap::new(),
        },
    );
    assert!(matches!(
        result,
        Err(Error::RowEditBlocked {
            reason: RowEditBlock::ProtectedSheet,
            ..
        })
    ));
}

#[test]
fn column_visibility_splits_effective_owners_and_preserves_other_properties() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:cols><x:col min="2" max="4" width="20" hidden="1" z:keep="yes"/><x:col min="3" max="3" width="10"/></x:cols><x:sheetData z:untouched="yes"/></x:worksheet>"#
    );
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::from([
                (Column::new(1).expect("B"), ColumnAction::show()),
                (Column::new(2).expect("C"), ColumnAction::hide()),
                (Column::new(4).expect("E"), ColumnAction::hide()),
            ]),
        },
    )
    .expect("column rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(r#"<x:col width="20" z:keep="yes" min="2" max="2"/>"#));
    assert!(text.contains(r#"<x:col width="20" hidden="1" z:keep="yes" min="3" max="4"/>"#));
    assert!(text.contains(r#"<x:col width="10" min="3" max="3" hidden="1"/>"#));
    assert!(text.contains(r#"<x:col min="5" max="5" hidden="1"/>"#));
    assert!(text.contains(r#"<x:sheetData z:untouched="yes"/>"#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse columns");
    let b = store.column(Column::new(1).expect("B"));
    assert!(!b.hidden());
    assert_eq!(b.width().map(Width::get), Some(20.0));
    let c = store.column(Column::new(2).expect("C"));
    assert!(c.hidden());
    assert_eq!(c.width().map(Width::get), Some(10.0));
    assert!(store.column(Column::new(3).expect("D")).hidden());
    assert!(store.column(Column::new(4).expect("E")).hidden());
}

#[test]
fn column_layout_facets_split_compactly_and_preserve_unedited_attributes() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:cols><x:col min="2" max="4" width="20" style="1" hidden="1" bestFit="1" customWidth="1" phonetic="1" outlineLevel="2" collapsed="1" z:keep="yes"/></x:cols><x:sheetData z:untouched="yes"/></x:worksheet>"#
    );
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::from([
                (
                    Column::new(1).expect("B"),
                    ColumnAction {
                        width: Some(WidthEffect::Reset),
                        best_fit: Some(false),
                        outline: Some(Outline::NONE),
                        collapsed: Some(false),
                        phonetic: Some(false),
                        ..ColumnAction::default()
                    },
                ),
                (
                    Column::new(2).expect("C"),
                    ColumnAction {
                        hidden: Some(false),
                        width: Some(WidthEffect::Set(Width::new(12.5).expect("width"))),
                        outline: Some(Outline::new(3).expect("outline")),
                        ..ColumnAction::default()
                    },
                ),
                (
                    Column::new(4).expect("E"),
                    ColumnAction {
                        width: Some(WidthEffect::Set(Width::new(15.0).expect("width"))),
                        best_fit: Some(true),
                        outline: Some(Outline::new(1).expect("outline")),
                        collapsed: Some(true),
                        phonetic: Some(true),
                        ..ColumnAction::default()
                    },
                ),
            ]),
        },
    )
    .expect("column layout rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(r#"style="1" hidden="1" z:keep="yes" min="2" max="2""#));
    assert!(text.contains(concat!(
        r#"style="1" bestFit="1" phonetic="1" collapsed="1" z:keep="yes" "#,
        r#"min="3" max="3" width="12.5" customWidth="1" outlineLevel="3""#
    )));
    assert!(text.contains(concat!(
        r#"<x:col min="5" max="5" width="15" customWidth="1" bestFit="1" "#,
        r#"outlineLevel="1" collapsed="1" phonetic="1"/>"#
    )));
    assert!(text.contains(r#"<x:sheetData z:untouched="yes"/>"#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse layout");
    let b = store.column(Column::new(1).expect("B"));
    assert_eq!(b.width(), None);
    assert!(b.hidden());
    assert!(!b.best_fit());
    assert_eq!(b.outline(), Outline::NONE);
    assert!(!b.collapsed());
    assert!(!b.phonetic());
    assert_eq!(
        store
            .column_entry(b.index())
            .map(|entry| entry.properties.style),
        Some(Some(1))
    );
    let c = store.column(Column::new(2).expect("C"));
    assert_eq!(c.width().map(Width::get), Some(12.5));
    assert!(!c.hidden());
    assert!(c.best_fit());
    assert_eq!(c.outline().get(), 3);
    assert!(c.collapsed());
    assert!(c.phonetic());
    let e = store.column(Column::new(4).expect("E"));
    assert_eq!(e.width().map(Width::get), Some(15.0));
    assert!(e.best_fit());
    assert_eq!(e.outline().get(), 1);
    assert!(e.collapsed());
    assert!(e.phonetic());
}

#[test]
fn column_style_retargeting_splits_ranges_and_resets_sparsely() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:cols><x:col min="2" max="4" width="20" style="1" z:keep="yes"/></x:cols><x:sheetData/></x:worksheet>"#
    );
    let edited = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::from([
                (
                    Column::new(1).expect("B"),
                    ColumnAction {
                        style: Some(StyleEffect::Reset),
                        ..ColumnAction::default()
                    },
                ),
                (
                    Column::new(2).expect("C"),
                    ColumnAction {
                        style: Some(StyleEffect::Set(2)),
                        ..ColumnAction::default()
                    },
                ),
                (
                    Column::new(4).expect("E"),
                    ColumnAction {
                        style: Some(StyleEffect::Set(3)),
                        width: Some(WidthEffect::Set(Width::new(12.0).expect("width"))),
                        ..ColumnAction::default()
                    },
                ),
                (
                    Column::new(5).expect("F"),
                    ColumnAction {
                        style: Some(StyleEffect::Reset),
                        ..ColumnAction::default()
                    },
                ),
            ]),
        },
    )
    .expect("column style rewrite");
    let text = std::str::from_utf8(&edited).expect("UTF-8");
    assert!(text.contains(r#"z:keep="yes""#));
    assert!(!text.contains(r#"min="6""#));

    let store = worksheet::parse(&edited, || Ok(None)).expect("reparse column styles");
    assert_eq!(
        store
            .column_entry(Column::new(1).expect("B"))
            .expect("B")
            .properties
            .style,
        None
    );
    assert_eq!(
        store
            .column_entry(Column::new(2).expect("C"))
            .expect("C")
            .properties
            .style,
        Some(2)
    );
    assert_eq!(
        store
            .column_entry(Column::new(3).expect("D"))
            .expect("D")
            .properties
            .style,
        Some(1)
    );
    assert_eq!(
        store
            .column_entry(Column::new(4).expect("E"))
            .expect("E")
            .properties
            .style,
        Some(3)
    );
    assert!(!store.column(Column::new(5).expect("F")).stored());
}

#[test]
fn style_only_implicit_column_is_blocked_before_zero_width_materialization() {
    let xml = format!(r#"<x:worksheet xmlns:x="{S}"><x:sheetData/></x:worksheet>"#);
    let result = rewrite(
        xml.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::from([(
                Column::new(2).expect("C"),
                ColumnAction {
                    style: Some(StyleEffect::Set(1)),
                    ..ColumnAction::default()
                },
            )]),
        },
    );
    assert!(matches!(
        result,
        Err(Error::ColumnEditBlocked {
            column,
            reason: ColumnEditBlock::StyleNeedsWidth,
            ..
        }) if column == Column::new(2).expect("C")
    ));
}

#[test]
fn column_visibility_inserts_sparse_cols_and_blocks_unsafe_splits() {
    let plain = format!(r#"<x:worksheet xmlns:x="{S}"><x:sheetData/></x:worksheet>"#);
    let inserted = rewrite(
        plain.as_bytes(),
        "Data",
        Plan {
            defaults: None,
            cells: BTreeMap::new(),
            rows: BTreeMap::new(),
            columns: BTreeMap::from([
                (Column::new(1).expect("B"), ColumnAction::hide()),
                (Column::new(2).expect("C"), ColumnAction::hide()),
            ]),
        },
    )
    .expect("insert cols");
    assert!(
        std::str::from_utf8(&inserted)
            .expect("UTF-8")
            .contains(r#"<x:cols><x:col min="2" max="3" hidden="1"/></x:cols><x:sheetData/>"#)
    );

    let extended = format!(
        r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><cols><col min="1" max="2"><z:future/></col></cols><sheetData/></worksheet>"#
    );
    assert!(matches!(
        rewrite(
            extended.as_bytes(),
            "Data",
            Plan {
                defaults: None,
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([(Column::new(0).expect("A"), ColumnAction::hide(),)]),
            },
        ),
        Err(Error::ColumnEditBlocked {
            reason: ColumnEditBlock::MarkupCompatibility,
            ..
        })
    ));

    let extended_columns = format!(
        r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><cols><col min="1" max="1"/><z:future/></cols><sheetData/></worksheet>"#
    );
    assert!(matches!(
        rewrite(
            extended_columns.as_bytes(),
            "Data",
            Plan {
                defaults: None,
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([(Column::new(1).expect("B"), ColumnAction::hide(),)]),
            },
        ),
        Err(Error::ColumnEditBlocked {
            reason: ColumnEditBlock::MarkupCompatibility,
            ..
        })
    ));

    let protected =
        format!(r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#);
    assert!(matches!(
        rewrite(
            protected.as_bytes(),
            "Data",
            Plan {
                defaults: None,
                cells: BTreeMap::new(),
                rows: BTreeMap::new(),
                columns: BTreeMap::from([(Column::new(0).expect("A"), ColumnAction::hide(),)]),
            },
        ),
        Err(Error::ColumnEditBlocked {
            reason: ColumnEditBlock::ProtectedSheet,
            ..
        })
    ));
}

#[test]
fn style_effects_preserve_payload_and_compose_with_value_effects() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:sheetData><x:row r="1"><x:c r="A1" s="1" z:keep="yes"><x:v>5</x:v></x:c><x:c r="B1" s="1"/></x:row></x:sheetData></x:worksheet>"#
    );
    let mut combined = Action::set(7_i32.into());
    combined.set_style(StyleEffect::Set(3));
    let actions = BTreeMap::from([
        (Address::from_a1("A1").unwrap(), Action::style(2)),
        (Address::from_a1("B1").unwrap(), Action::reset_style()),
        (Address::from_a1("C1").unwrap(), Action::style(3)),
        (Address::from_a1("D1").unwrap(), combined),
    ]);

    let edited = rewrite(xml.as_bytes(), "Data", actions).unwrap();
    let edited = std::str::from_utf8(&edited).unwrap();
    assert!(edited.contains(r#"z:keep="yes" r="A1" s="2"><x:v>5</x:v>"#));
    assert!(edited.contains(r#"<x:c r="B1"/>"#));
    assert!(edited.contains(r#"<x:c r="C1" s="3"/>"#));
    assert!(edited.contains(r#"<x:c r="D1" s="3"><x:v>7</x:v></x:c>"#));

    let store = worksheet::parse(edited.as_bytes(), || Ok(None)).unwrap();
    assert_eq!(
        store.entry(Address::from_a1("A1").unwrap()).unwrap().style,
        Some(2)
    );
    assert_eq!(
        store.entry(Address::from_a1("B1").unwrap()).unwrap().style,
        None
    );
    assert!(matches!(
        store.get(Address::from_a1("C1").unwrap()),
        Some(Cell::Empty)
    ));
}

#[test]
fn merge_surgery_is_lossless_ordered_and_dependency_checked() {
    let xml = format!(
        r#"<x:worksheet xmlns:x="{S}" xmlns:z="urn:future"><x:dimension z:keep="dimension" ref="A1"/><x:sheetData/><x:mergeCells z:keep="container" count="1"><x:mergeCell z:keep="record" ref="E5:F5"/></x:mergeCells><x:hyperlinks/></x:worksheet>"#
    );
    let added = rewrite_merges(
        xml.as_bytes(),
        "Data",
        MergePlan {
            add: vec![Rect::from_a1("B2:C3").expect("range")],
            remove: Vec::new(),
        },
    )
    .expect("add merge");
    let added_text = std::str::from_utf8(&added).expect("UTF-8");
    assert!(added_text.contains(r#"z:keep="dimension" ref="A1:C3""#));
    assert!(added_text.contains(r#"z:keep="container" count="2""#));
    assert!(added_text.contains(r#"<x:mergeCell z:keep="record" ref="E5:F5"/>"#));
    assert!(added_text.contains(r#"<x:mergeCell ref="B2:C3"/>"#));
    assert!(
        added_text.find("<x:mergeCells").expect("merge container")
            < added_text.find("<x:hyperlinks").expect("successor")
    );

    let removed = rewrite_merges(
        &added,
        "Data",
        MergePlan {
            add: Vec::new(),
            remove: vec![Rect::from_a1("E5:F5").expect("range")],
        },
    )
    .expect("remove merge");
    let removed_text = std::str::from_utf8(&removed).expect("UTF-8");
    assert!(!removed_text.contains("E5:F5"));
    assert!(removed_text.contains(r#"z:keep="container" count="1""#));

    let emptied = rewrite_merges(
        &removed,
        "Data",
        MergePlan {
            add: Vec::new(),
            remove: vec![Rect::from_a1("B2:C3").expect("range")],
        },
    )
    .expect("remove final merge");
    let emptied = std::str::from_utf8(&emptied).expect("UTF-8");
    assert!(!emptied.contains("mergeCells"));
    assert!(
        emptied.contains(r#"ref="A1:C3""#),
        "dimensions never shrink"
    );

    let requested = Rect::from_a1("A1:B2").expect("range");
    for (xml, expected) in [
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData/><sheetProtection sheet="1"/></worksheet>"#
            ),
            MergeEditBlock::ProtectedSheet,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B2">A1:B2</f></c></row></sheetData></worksheet>"#
            ),
            MergeEditBlock::GroupFormula,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><sheetData/><mergeCells><mergeCell ref="C3:D4"/><z:future/></mergeCells></worksheet>"#
            ),
            MergeEditBlock::UnmodeledPayload,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><sheetData/><mc:AlternateContent/></worksheet>"#
            ),
            MergeEditBlock::MarkupCompatibility,
        ),
    ] {
        assert!(matches!(
            rewrite_merges(
                xml.as_bytes(),
                "Data",
                MergePlan {
                    add: vec![requested],
                    remove: Vec::new(),
                },
            ),
            Err(Error::MergeEditBlocked { reason, .. }) if reason == expected
        ));
    }
}

#[test]
fn blocks_dependencies_instead_of_guessing() {
    let cases = [
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><sheetProtection sheet="1"/></worksheet>"#
            ),
            "A1",
            EditBlock::ProtectedSheet,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><mergeCells><mergeCell ref="A1:B2"/></mergeCells></worksheet>"#
            ),
            "B2",
            EditBlock::CoveredMerge,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"><f t="array" ref="A1:B2">A1:B2*2</f></c></row></sheetData></worksheet>"#
            ),
            "B2",
            EditBlock::GroupFormula,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}"><sheetData><row r="1"><c r="A1"/></row></sheetData><dataValidations count="1"><dataValidation sqref="A1:B2"/></dataValidations></worksheet>"#
            ),
            "B2",
            EditBlock::DataValidation,
        ),
        (
            format!(
                r#"<worksheet xmlns="{S}" xmlns:z="urn:future"><sheetData><row r="1"><c r="A1"><z:value/></c></row></sheetData></worksheet>"#
            ),
            "A1",
            EditBlock::MarkupCompatibility,
        ),
    ];
    for (xml, address, expected) in cases {
        let address = Address::from_a1(address).unwrap();
        let actions = BTreeMap::from([(address, Action::set(1_i32.into()))]);
        assert!(matches!(
            rewrite(xml.as_bytes(), "Data", actions),
            Err(Error::EditBlocked { reason, .. }) if reason == expected
        ));
    }
}
