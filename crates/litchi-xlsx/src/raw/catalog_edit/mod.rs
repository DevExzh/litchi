//! Narrow, lossless surgery for the workbook sheet catalog.
//!
//! The facade keeps the existing raw catalog-edit surface stable while the
//! implementation is divided into semantic models, XML codecs, validation,
//! and atomic package orchestration.
mod codec;
mod model;
mod package;
mod validation;

pub(crate) use codec::dialect;
pub(crate) use model::{
    Active, Create, Dialect, MAX_ACTIVE_TAB, MAX_SHEET_ID, MAX_SHEETS, Order, Plan, Remove, Rename,
    State, Tab,
};
pub(crate) use package::{append, remove, replace_defined_names, rewrite};

#[cfg(test)]
mod tests {
    use super::codec::MCE;
    use super::model::FIRST_SHEET_SENTINEL;
    use super::*;
    use crate::error::{Error, TabEditBlock};
    use crate::raw::{Visibility, parse_catalog};

    const S: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn plan<'a>(tabs: Vec<Tab<'a>>, active: Option<usize>) -> Plan<'a> {
        Plan {
            tabs,
            renames: Vec::new(),
            active: active.map(|position| Active {
                sheet: "Active",
                position,
            }),
            order: None,
        }
    }

    #[test]
    fn rewrites_only_selected_states_and_first_view() {
        let source = format!(
            r#"<?xml version="1.0"?><x:workbook xmlns:x="{S}" xmlns:rel="{R}" xmlns:z="urn:future"><x:bookViews><x:workbookView activeTab="0" z:keep="yes"/><x:workbookView activeTab="0"/></x:bookViews><x:sheets z:container="exact"><x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes"/><x:sheet name="Two" sheetId="2" state="hidden" rel:id="r2"/></x:sheets><x:extLst><z:data value="exact"/></x:extLst></x:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![
                    Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    },
                    Tab {
                        sheet: "Two",
                        position: 1,
                        relationship_id: "r2",
                        state: State::Visible,
                    },
                ],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<x:sheet name="One" sheetId="1" rel:id="r1" z:keep="yes" state="hidden"/>"#
        ));
        assert!(text.contains(r#"<x:sheet name="Two" sheetId="2" rel:id="r2"/>"#));
        assert!(text.contains(r#"<x:workbookView z:keep="yes" activeTab="1"/>"#));
        assert!(text.contains(r#"<x:workbookView activeTab="0"/>"#));
        assert!(text.contains(r#"<x:extLst><z:data value="exact"/></x:extLst>"#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.active_sheet_index, 1);
        assert_eq!(catalog.sheets[0].visibility, Visibility::Hidden);
        assert_eq!(catalog.sheets[1].visibility, Visibility::Visible);
    }

    #[test]
    fn composes_name_and_visibility_on_one_lossless_sheet_slot() {
        let source = format!(
            r#"<x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future"><x:sheets><x:sheet name="Data" sheetId="7" r:id="r1" z:keep="exact"/></x:sheets></x:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            Plan {
                tabs: vec![Tab {
                    sheet: "Data",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                renames: vec![Rename {
                    sheet: "Data",
                    position: 0,
                    relationship_id: "r1",
                    name: "Input 2026",
                }],
                active: None,
                order: None,
            },
        )
        .expect("catalog rewrite");
        assert_eq!(
            std::str::from_utf8(&output).expect("UTF-8"),
            format!(
                r#"<x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future"><x:sheets><x:sheet sheetId="7" r:id="r1" z:keep="exact" state="hidden" name="Input 2026"/></x:sheets></x:workbook>"#
            )
        );
    }

    #[test]
    fn inserts_prefixed_book_views_before_sheets() {
        let source = format!(
            r#"<s:workbook xmlns:s="{S}" xmlns:r="{R}"><s:sheets><s:sheet name="One" sheetId="1" state="hidden" r:id="r1"/><s:sheet name="Two" sheetId="2" r:id="r2"/></s:sheets></s:workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(
            text.contains(
                r#"<s:bookViews><s:workbookView activeTab="1"/></s:bookViews><s:sheets>"#
            )
        );
        assert_eq!(
            parse_catalog(&output).expect("catalog").active_sheet_index,
            1
        );
    }

    #[test]
    fn expands_an_empty_book_views_container() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews data="kept"/><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        let output = rewrite(
            source.as_bytes(),
            plan(
                vec![Tab {
                    sheet: "One",
                    position: 0,
                    relationship_id: "r1",
                    state: State::Hidden,
                }],
                Some(1),
            ),
        )
        .expect("rewrite");
        assert!(
            std::str::from_utf8(&output)
                .expect("UTF-8")
                .contains(r#"<bookViews data="kept"><workbookView activeTab="1"/></bookViews>"#)
        );
    }

    #[test]
    fn blocks_structure_protection_but_preserves_unrelated_alternate_content() {
        let protected = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><workbookProtection lockStructure="1"/><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        let tab = Tab {
            sheet: "One",
            position: 0,
            relationship_id: "r1",
            state: State::Hidden,
        };
        assert!(matches!(
            rewrite(protected.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
        let activated = rewrite(protected.as_bytes(), plan(Vec::new(), Some(0)))
            .expect("structure protection permits active-tab selection");
        assert_eq!(
            parse_catalog(&activated)
                .expect("protected catalog")
                .active_sheet_index,
            0
        );

        let mce = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        let rewritten =
            rewrite(mce.as_bytes(), plan(vec![tab], None)).expect("unrelated compatibility XML");
        let text = std::str::from_utf8(&rewritten).expect("UTF-8");
        assert!(text.contains("mc:AlternateContent"));
        assert!(text.contains(r#"state="hidden""#));

        let nested_protection = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback><workbookProtection lockStructure="1"/></mc:Fallback></mc:AlternateContent><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(nested_protection.as_bytes(), plan(vec![tab], None)),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ProtectedWorkbook,
                ..
            })
        ));
    }

    #[test]
    fn blocks_active_view_insertion_beside_alternate_content() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}"><mc:AlternateContent><mc:Fallback/></mc:AlternateContent><sheets><sheet name="One" sheetId="1" state="hidden" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "One",
                        position: 0,
                        relationship_id: "r1",
                        state: State::Hidden,
                    }],
                    Some(1),
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
        assert!(matches!(
            rewrite(source.as_bytes(), plan(Vec::new(), Some(1))),
            Err(Error::TabEditBlocked {
                sheet,
                position: 1,
                reason: TabEditBlock::MarkupCompatibility,
            }) if sheet == "Active"
        ));
    }

    #[test]
    fn blocks_an_effective_sheet_without_a_direct_slot() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                plan(
                    vec![Tab {
                        sheet: "Fallback",
                        position: 1,
                        relationship_id: "mce-rel",
                        state: State::VeryHidden,
                    }],
                    None,
                )
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }

    #[test]
    fn active_tab_limit_is_a_typed_block() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                source.as_bytes(),
                Plan {
                    tabs: Vec::new(),
                    renames: Vec::new(),
                    active: Some(Active {
                        sheet: "Too Far",
                        position: MAX_ACTIVE_TAB + 1,
                    }),
                    order: None,
                }
            ),
            Err(Error::TabEditBlocked {
                sheet,
                position,
                reason: TabEditBlock::ActiveTabLimit,
            }) if sheet == "Too Far" && position == MAX_ACTIVE_TAB + 1
        ));
    }

    #[test]
    fn reorders_losslessly_and_remaps_every_positional_dependency() {
        let source = format!(
            r#"<?xml version="1.0"?><x:workbook xmlns:x="{S}" xmlns:r="{R}" xmlns:z="urn:future" xmlns:mc="{mce}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15"><mc:AlternateContent><mc:Choice Requires="x15"><x15ac:absPath url="/exact/" xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"/></mc:Choice></mc:AlternateContent><x:bookViews><x:workbookView activeTab="2" firstSheet="1" z:keep="view-one"/><x:workbookView activeTab="0" firstSheet="{FIRST_SHEET_SENTINEL}" z:keep="view-two"/><x:workbookView z:keep="defaults"/></x:bookViews><x:sheets><x:sheet name="One" sheetId="10" r:id="r1" z:keep="one"/><x:sheet name="Two" sheetId="20" r:id="r2"/><x:sheet name="Three" sheetId="30" state="hidden" r:id="r3"/></x:sheets><x:definedNames><x:definedName name="OneLocal" localSheetId="0">One!$A$1</x:definedName><x:definedName name="ThreeLocal" localSheetId="2">Three!$A$1</x:definedName><x:definedName name="Global">1</x:definedName></x:definedNames><x:customWorkbookViews><x:customWorkbookView name="Exact" guid="{{00000000-0000-0000-0000-000000000001}}" activeSheetId="30"/></x:customWorkbookViews></x:workbook>"#,
            mce = String::from_utf8_lossy(MCE)
        );
        let output = rewrite(
            source.as_bytes(),
            Plan {
                tabs: vec![Tab {
                    sheet: "Two",
                    position: 1,
                    relationship_id: "r2",
                    state: State::VeryHidden,
                }],
                renames: Vec::new(),
                active: None,
                order: Some(Order {
                    sheet: "Three",
                    position: 2,
                    relationship_ids: vec!["r3", "r1", "r2"],
                    local_scopes: 2,
                }),
            },
        )
        .expect("reorder");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        let three = text.find("name=\"Three\"").expect("Three");
        let one = text.find("name=\"One\"").expect("One");
        let two = text.find("name=\"Two\"").expect("Two");
        assert!(three < one && one < two);
        assert!(text.contains(r#"name="One" sheetId="10" r:id="r1" z:keep="one""#));
        assert!(text.contains(r#"name="Two" sheetId="20" r:id="r2" state="veryHidden""#));
        assert!(text.contains(r#"z:keep="view-one" activeTab="0" firstSheet="2""#));
        assert!(text.contains(&format!(
            r#"firstSheet="{FIRST_SHEET_SENTINEL}" z:keep="view-two" activeTab="1""#
        )));
        assert!(text.contains(r#"z:keep="defaults" activeTab="1" firstSheet="1""#));
        assert!(text.contains(r#"name="OneLocal" localSheetId="1""#));
        assert!(text.contains(r#"name="ThreeLocal" localSheetId="0""#));
        assert!(text.contains(r#"activeSheetId="30""#));
        assert!(text.contains(r#"<x15ac:absPath url="/exact/""#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.active_sheet_index, 0);
        assert_eq!(
            catalog
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["Three", "One", "Two"]
        );
        assert_eq!(catalog.defined_names[0].local_sheet_id, Some(1));
        assert_eq!(catalog.defined_names[1].local_sheet_id, Some(0));
    }

    #[test]
    fn reorder_blocks_unmodeled_catalogs_and_invalid_secondary_views() {
        let order = || Order {
            sheet: "Two",
            position: 1,
            relationship_ids: vec!["r2", "r1"],
            local_scopes: 0,
        };
        for source in [
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><future/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView/><future/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets><definedNames><future/></definedNames></workbook>"#
            ),
            format!(
                r#"<workbook xmlns="{S}" xmlns:r="{R}" xmlns:mc="{mce}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" mc:Ignorable="x15"><mc:AlternateContent><mc:Choice Requires="x15"><bookViews><workbookView activeTab="1"/></bookViews></mc:Choice></mc:AlternateContent><bookViews><workbookView/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#,
                mce = String::from_utf8_lossy(MCE)
            ),
        ] {
            assert!(matches!(
                rewrite(
                    source.as_bytes(),
                    Plan {
                        tabs: Vec::new(),
                        renames: Vec::new(),
                        active: None,
                        order: Some(order()),
                    }
                ),
                Err(Error::TabEditBlocked {
                    reason: TabEditBlock::MarkupCompatibility,
                    ..
                })
            ));
        }

        let invalid_view = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><bookViews><workbookView/><workbookView activeTab="9"/></bookViews><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        assert!(matches!(
            rewrite(
                invalid_view.as_bytes(),
                Plan {
                    tabs: Vec::new(),
                    renames: Vec::new(),
                    active: None,
                    order: Some(order()),
                }
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::ViewIndex,
                ..
            })
        ));
    }

    #[test]
    fn append_preserves_existing_bytes_and_uses_the_document_prefixes() {
        let source = format!(
            r#"<?xml version="1.0"?><s:workbook xmlns:s="{S}" xmlns:rel="{R}" xmlns:x="urn:keep" x:exact="yes"><s:bookViews><s:workbookView activeTab="0"/></s:bookViews><s:sheets><s:sheet name="One" sheetId="7" rel:id="tab" x:keep="1"/></s:sheets><x:tail>opaque</x:tail></s:workbook>"#
        );
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "A&B",
                position: 1,
                sheet_id: 1,
                relationship_id: "rId1",
                state: State::Hidden,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(
            r#"<s:sheet name="One" sheetId="7" rel:id="tab" x:keep="1"/><s:sheet name="A&amp;B" sheetId="1" rel:id="rId1" state="hidden"/>"#
        ));
        assert!(text.contains(r#"<x:tail>opaque</x:tail>"#));
        assert!(text.contains(r#"<s:workbook xmlns:s="#));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(catalog.sheets.len(), 2);
        assert_eq!(catalog.sheets[1].name, "A&B");
        assert!(matches!(catalog.sheets[1].visibility, Visibility::Hidden));
    }

    #[test]
    fn append_expands_an_empty_sheet_container_from_root_namespaces() {
        let source =
            format!(r#"<s:workbook xmlns:s="{S}" xmlns:rel="{R}"><s:sheets/></s:workbook>"#);
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "Only",
                position: 0,
                sheet_id: 9,
                relationship_id: "new",
                state: State::Visible,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(
            text.contains(
                r#"<s:sheets><s:sheet name="Only" sheetId="9" rel:id="new"/></s:sheets>"#
            )
        );
        assert_eq!(parse_catalog(&output).expect("catalog").sheets.len(), 1);
    }

    #[test]
    fn append_expands_an_empty_sheet_container_with_a_local_relationship_prefix() {
        let source =
            format!(r#"<s:workbook xmlns:s="{S}"><s:sheets xmlns:rel="{R}"/></s:workbook>"#);
        let output = append(
            source.as_bytes(),
            Create {
                sheet: "Only",
                position: 0,
                sheet_id: 1,
                relationship_id: "rId1",
                state: State::Visible,
            },
        )
        .expect("append");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(text.contains(r#"<s:sheet name="Only" sheetId="1" rel:id="rId1"/>"#));
        assert_eq!(parse_catalog(&output).expect("catalog").sheets.len(), 1);
    }

    #[test]
    fn removal_drops_slots_and_scopes_and_remaps_views_losslessly() {
        let source = format!(
            r#"<?xml version="1.0"?><s:workbook xmlns:s="{S}" xmlns:r="{R}" xmlns:k="urn:keep" k:root="exact"><s:bookViews><s:workbookView activeTab="1" firstSheet="1" k:v="first"/><s:workbookView activeTab="2" firstSheet="2" k:v="second"/></s:bookViews><s:sheets><s:sheet name="One" sheetId="7" r:id="r1" k:x="one"/><s:sheet name="Middle" sheetId="9" r:id="r2"/><s:sheet name="Three" sheetId="11" r:id="r3"/></s:sheets><s:definedNames><s:definedName name="OneLocal" localSheetId="0">One!A1</s:definedName><s:definedName name="MiddleLocal" localSheetId="1">Middle!A1</s:definedName><s:definedName name="ThreeLocal" localSheetId="2">Three!A1</s:definedName><s:definedName name="Global">1</s:definedName></s:definedNames><k:tail>opaque</k:tail></s:workbook>"#
        );
        let output = remove(
            source.as_bytes(),
            Remove {
                sheet: "Middle",
                position: 1,
                relationship_ids: vec!["r2"],
                active: Active {
                    sheet: "Three",
                    position: 1,
                },
                local_scopes: 3,
            },
        )
        .expect("remove");
        let text = std::str::from_utf8(&output).expect("UTF-8");
        assert!(!text.contains("name=\"Middle\""));
        assert!(!text.contains("name=\"MiddleLocal\""));
        assert!(text.contains(r#"name="One" sheetId="7" r:id="r1" k:x="one""#));
        assert!(text.contains(r#"activeTab="1" firstSheet="1" k:v="first""#));
        assert!(text.contains(r#"k:v="second" activeTab="1" firstSheet="1""#));
        assert!(text.contains(r#"name="ThreeLocal" localSheetId="1""#));
        assert!(text.contains("<k:tail>opaque</k:tail>"));
        let catalog = parse_catalog(&output).expect("catalog");
        assert_eq!(
            catalog
                .sheets
                .iter()
                .map(|sheet| sheet.name.as_str())
                .collect::<Vec<_>>(),
            ["One", "Three"]
        );
        assert_eq!(catalog.active_sheet_index, 1);
        assert_eq!(catalog.defined_names.len(), 3);
        assert_eq!(catalog.defined_names[1].local_sheet_id, Some(1));
    }

    #[test]
    fn removal_drops_an_empty_defined_name_container() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><sheet name="Two" sheetId="2" r:id="r2"/></sheets><definedNames><definedName name="Only" localSheetId="1">Two!A1</definedName></definedNames></workbook>"#
        );
        let output = remove(
            source.as_bytes(),
            Remove {
                sheet: "Two",
                position: 1,
                relationship_ids: vec!["r2"],
                active: Active {
                    sheet: "One",
                    position: 0,
                },
                local_scopes: 1,
            },
        )
        .expect("remove");
        assert!(
            !std::str::from_utf8(&output)
                .expect("UTF-8")
                .contains("definedNames")
        );
    }

    #[test]
    fn removal_blocks_unmodeled_catalog_payloads() {
        let source = format!(
            r#"<workbook xmlns="{S}" xmlns:r="{R}"><sheets><sheet name="One" sheetId="1" r:id="r1"/><future/><sheet name="Two" sheetId="2" r:id="r2"/></sheets></workbook>"#
        );
        assert!(matches!(
            remove(
                source.as_bytes(),
                Remove {
                    sheet: "Two",
                    position: 1,
                    relationship_ids: vec!["r2"],
                    active: Active {
                        sheet: "One",
                        position: 0,
                    },
                    local_scopes: 0,
                }
            ),
            Err(Error::TabEditBlocked {
                reason: TabEditBlock::MarkupCompatibility,
                ..
            })
        ));
    }
}
