use super::codec::{self, OwnerKind};
use super::{SHOW_DATA_TYPE_ICONS_NAMESPACE, ShowDataTypeIcons, Target};
use crate::named_sheet_view::{
    Extension, Guid, View, Views, parse_named_sheet_views, store_worksheet_named_sheet_views,
};
use crate::sheet_view::parse_worksheet_views;
use litchi_opc::{OpcPackage, PackURI};
use std::sync::Arc;

const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const NSV: &str = "http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews";

#[test]
fn reads_both_elements_with_inherited_namespace_bindings() {
    let worksheet = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="urn:producer"><sdt:showDataTypeIcons visible="0"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    let views = parse_worksheet_views(worksheet.as_bytes())
        .unwrap()
        .unwrap();
    assert!(!views.entries()[0].show_data_type_icons().unwrap().visible());

    let named = format!(
        r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="urn:producer"><sdt:showDataTypeIconsCustomSheetView/></x:ext></extLst></namedSheetView></namedSheetViews>"#
    );
    let views = parse_named_sheet_views(named.as_bytes()).unwrap();
    assert!(
        views.views()[0]
            .show_data_type_icons_custom_sheet_view()
            .unwrap()
            .visible()
    );
}

#[test]
fn rejects_duplicates_misplaced_payloads_and_bad_booleans() {
    let duplicate = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons/><sdt:showDataTypeIcons/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    assert!(parse_worksheet_views(duplicate.as_bytes()).is_err());
    let wrong_owner = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><extLst><ext uri="a"><sdt:showDataTypeIcons/></ext></extLst><sheetView workbookViewId="0"/></sheetViews></worksheet>"#
    );
    assert!(parse_worksheet_views(wrong_owner.as_bytes()).is_err());
    let bad_boolean = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons visible="yes"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    assert!(parse_worksheet_views(bad_boolean.as_bytes()).is_err());
    let nested = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" xmlns:u="urn:unknown"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><u:wrapper><sdt:showDataTypeIcons/></u:wrapper></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    assert!(parse_worksheet_views(nested.as_bytes()).is_err());
}

#[test]
fn rejects_entity_content_and_accepts_schema_collapsed_boolean_whitespace() {
    let spaced = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons visible="  true  "/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    assert!(
        parse_worksheet_views(spaced.as_bytes())
            .unwrap()
            .unwrap()
            .entries()[0]
            .show_data_type_icons()
            .unwrap()
            .visible()
    );

    for content in ["&amp;", "&#x20;"] {
        let payload = format!(
            r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons>{content}</sdt:showDataTypeIcons></ext></extLst></sheetView></sheetViews></worksheet>"#
        );
        assert!(
            parse_worksheet_views(payload.as_bytes()).is_err(),
            "{content}"
        );
        assert!(
            codec::rewrite(
                payload.as_bytes(),
                OwnerKind::WorksheetView,
                0,
                Some(ShowDataTypeIcons::new(false)),
            )
            .is_err(),
            "source rewrite accepted {content}"
        );

        let named = format!(
            r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="a"><sdt:showDataTypeIconsCustomSheetView>{content}</sdt:showDataTypeIconsCustomSheetView></x:ext></extLst></namedSheetView></namedSheetViews>"#
        );
        assert!(
            parse_named_sheet_views(named.as_bytes()).is_err(),
            "{content}"
        );
    }

    for payload in [
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons> </sdt:showDataTypeIcons></ext></extLst></sheetView></sheetViews></worksheet>"#
        ),
        format!(
            r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons><![CDATA[ ]]></sdt:showDataTypeIcons></ext></extLst></sheetView></sheetViews></worksheet>"#
        ),
    ] {
        assert!(parse_worksheet_views(payload.as_bytes()).is_err());
        assert!(
            codec::rewrite(
                payload.as_bytes(),
                OwnerKind::WorksheetView,
                0,
                Some(ShowDataTypeIcons::new(false)),
            )
            .is_err()
        );
    }

    for payload in [
        format!(
            r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="a"><sdt:showDataTypeIconsCustomSheetView> </sdt:showDataTypeIconsCustomSheetView></x:ext></extLst></namedSheetView></namedSheetViews>"#
        ),
        format!(
            r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="a"><sdt:showDataTypeIconsCustomSheetView><![CDATA[ ]]></sdt:showDataTypeIconsCustomSheetView></x:ext></extLst></namedSheetView></namedSheetViews>"#
        ),
    ] {
        assert!(parse_named_sheet_views(payload.as_bytes()).is_err());
        assert!(
            codec::rewrite(
                payload.as_bytes(),
                OwnerKind::CustomSheetView,
                0,
                Some(ShowDataTypeIcons::new(false)),
            )
            .is_err()
        );
    }

    let authored = format!(
        r#"<x:ext xmlns:x="urn:producer" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sdt:showDataTypeIcons> </sdt:showDataTypeIcons></x:ext>"#
    );
    assert!(codec::inspect_extension(authored.as_bytes(), Target::Worksheet).is_err());
    assert!(
        codec::rewrite_extension(
            authored.as_bytes(),
            Target::Worksheet,
            Some(ShowDataTypeIcons::new(false)),
        )
        .is_err()
    );

    let nbsp = '\u{00a0}';
    let worksheet = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="a"><sdt:showDataTypeIcons visible="{nbsp}true{nbsp}"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    assert!(parse_worksheet_views(worksheet.as_bytes()).is_err());
    let named = format!(
        r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="a"><sdt:showDataTypeIconsCustomSheetView visible="{nbsp}true{nbsp}"/></x:ext></extLst></namedSheetView></namedSheetViews>"#
    );
    assert!(parse_named_sheet_views(named.as_bytes()).is_err());
}

#[test]
fn retains_ignorable_data_type_icon_extensions_in_both_owner_contexts() {
    let worksheet = format!(
        r#"<worksheet xmlns="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" mc:Ignorable="sdt"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="urn:producer"><sdt:showDataTypeIcons visible="0"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    let views = parse_worksheet_views(worksheet.as_bytes())
        .unwrap()
        .unwrap();
    assert!(!views.entries()[0].show_data_type_icons().unwrap().visible());

    let named = format!(
        r#"<namedSheetViews xmlns="{NSV}" xmlns:x="{SML}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" mc:Ignorable="sdt"><namedSheetView name="Data" id="{{01234567-89AB-CDEF-0123-456789ABCDEF}}"><extLst><x:ext uri="urn:producer"><sdt:showDataTypeIconsCustomSheetView visible="0"/></x:ext></extLst></namedSheetView></namedSheetViews>"#
    );
    let views = parse_named_sheet_views(named.as_bytes()).unwrap();
    assert!(
        !views.views()[0]
            .show_data_type_icons_custom_sheet_view()
            .unwrap()
            .visible()
    );
}

#[test]
fn rewrites_one_owner_and_preserves_unknown_extension_bytes() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" xmlns:u="urn:unknown"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="urn:producer"><u:keep a="1"/><sdt:showDataTypeIcons visible="0"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    let changed = codec::rewrite(
        xml.as_bytes(),
        OwnerKind::WorksheetView,
        0,
        Some(ShowDataTypeIcons::new(true)),
    )
    .unwrap();
    let changed = std::str::from_utf8(&changed).unwrap();
    assert!(changed.contains(r#"<u:keep a="1"/>"#));
    assert!(changed.contains("showDataTypeIcons"));
    assert!(changed.contains(r#"visible="1""#));
    let removed = codec::rewrite(changed.as_bytes(), OwnerKind::WorksheetView, 0, None).unwrap();
    assert!(
        !std::str::from_utf8(&removed)
            .unwrap()
            .contains("showDataTypeIcons visible")
    );
}

#[test]
fn rewrites_the_selected_view_when_multiple_owners_exist() {
    let xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="one"><sdt:showDataTypeIcons visible="0"/></ext></extLst></sheetView><sheetView workbookViewId="1"><extLst><ext uri="two"><sdt:showDataTypeIcons visible="1"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    let changed = codec::rewrite(
        xml.as_bytes(),
        OwnerKind::WorksheetView,
        1,
        Some(ShowDataTypeIcons::new(false)),
    )
    .unwrap();
    let changed = std::str::from_utf8(&changed).unwrap();
    assert!(changed.contains(r#"uri="one"><sdt:showDataTypeIcons visible="0""#));
    assert!(changed.contains(r#"uri="two"><sdt:showDataTypeIcons"#));
    assert!(changed.contains(r#"uri="two"><sdt:showDataTypeIcons xmlns:sdt="#));
}

fn package_with_icon_owner() -> (OpcPackage, PackURI) {
    let bytes =
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx");
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"><extLst><ext uri="urn:producer"><u:keep xmlns:u="urn:unknown"/></ext></extLst></sheetView></sheetViews></worksheet>"#
    );
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .set_blob(xml.into_bytes());
    (package, worksheet)
}

#[test]
fn adding_without_an_existing_extension_owner_is_explicitly_refused() {
    let bytes =
        include_bytes!("../../../../test-data/poi/test-data/spreadsheet/right-to-left.xlsx");
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let worksheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
    let xml = format!(
        r#"<worksheet xmlns="{SML}" xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}"><sheetViews><sheetView workbookViewId="0"/></sheetViews></worksheet>"#
    );
    package
        .get_part_mut(&worksheet)
        .unwrap()
        .set_blob(xml.into_bytes());
    let original = package.get_part(&worksheet).unwrap().blob().to_vec();
    let mut edit = super::edit(&mut package, &worksheet).unwrap();
    edit.set_worksheet_view(0, Some(ShowDataTypeIcons::new(false)))
        .unwrap();
    let error = edit.commit().unwrap_err();
    assert!(error.to_string().contains("existing ext owner"));
    assert_eq!(package.get_part(&worksheet).unwrap().blob(), original);
}

#[test]
fn source_transaction_has_exact_noop_and_inverse_restore() {
    let (mut package, worksheet) = package_with_icon_owner();
    let before = super::load(&package, &worksheet).unwrap();
    let source = before.source_arc();
    let mut noop = super::edit(&mut package, &worksheet).unwrap();
    assert!(!noop.set_worksheet_view(0, None).unwrap());
    let noop_commit = noop.commit().unwrap();
    assert!(!noop_commit.changed());
    assert!(noop_commit.patch().is_empty());
    assert!(Arc::ptr_eq(&source, &noop_commit.snapshot().source_arc()));

    let mut edit = super::edit(&mut package, &worksheet).unwrap();
    assert!(
        edit.set_worksheet_view(0, Some(ShowDataTypeIcons::new(false)))
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert!(!commit.patch().is_empty());
    assert!(!commit.snapshot().worksheet_views()[0].unwrap().visible());
    let changed_xml = package.get_part(&worksheet).unwrap().blob().to_vec();
    assert!(
        std::str::from_utf8(&changed_xml)
            .unwrap()
            .contains("showDataTypeIcons")
    );
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(
        package.get_part(&worksheet).unwrap().blob(),
        before.source_xml()
    );
}

#[test]
fn source_transaction_edits_custom_sheet_view_without_rebuilding_other_views() {
    let (mut package, worksheet) = package_with_icon_owner();
    let mut view = View::with_id(
        "Data",
        Guid::new("{01234567-89AB-CDEF-0123-456789ABCDEF}").unwrap(),
    )
    .unwrap();
    view.add_extension(Extension::new(
        "urn:producer",
        format!(
            r#"<sdt:showDataTypeIconsCustomSheetView xmlns:sdt="{SHOW_DATA_TYPE_ICONS_NAMESPACE}" visible="1"/>"#
        ),
    ).unwrap()).unwrap();
    store_worksheet_named_sheet_views(&mut package, &worksheet, &Views::new(view)).unwrap();
    let before = super::load(&package, &worksheet).unwrap();
    assert!(before.custom_sheet_views()[0].unwrap().visible());
    let named_before = before.named_sheet_views_source_arc().unwrap();

    let mut edit = super::edit(&mut package, &worksheet).unwrap();
    assert!(
        edit.set_custom_sheet_view(0, Some(ShowDataTypeIcons::new(false)))
            .unwrap()
    );
    let commit = edit.commit().unwrap();
    assert!(!commit.snapshot().custom_sheet_views()[0].unwrap().visible());
    let named_after = commit.snapshot().named_sheet_views_source_arc().unwrap();
    assert!(!Arc::ptr_eq(&named_before, &named_after));
    let named_xml = std::str::from_utf8(named_after.as_slice()).unwrap();
    assert!(named_xml.contains("showDataTypeIconsCustomSheetView"));
    assert!(named_xml.contains(r#"visible="0""#));
    commit.patch().inverse().apply(&mut package).unwrap();
    assert_eq!(
        package
            .get_part(commit.patch().before().named_sheet_views_part().unwrap())
            .unwrap()
            .blob(),
        named_before.as_slice()
    );
}
