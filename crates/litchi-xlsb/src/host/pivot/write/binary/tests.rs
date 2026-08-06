//! Focused PivotCache binary-writer tests.

use super::semantic::write_pivot_cache_definition;
use crate::package::error::Error;
use crate::package::pivot::model::*;
use crate::package::pivot::parse_pivot_cache_definition;
use crate::writer::{MutableWorksheet, WorkbookWriter};
use std::io::Cursor;

fn field(name: &str) -> PivotCacheField {
    PivotCacheField {
        name: name.to_string(),
        caption: None,
        number_format: None,
        sql_type: 0,
        hierarchy_index: 0,
        level: 0x7FFF,
        member_property_fields: Vec::new(),
        member_property_name: None,
        formula: None,
        server_based: false,
        cant_get_unique_items: false,
        source_field: true,
        olap_member_property_field: false,
        ignore: false,
        shared_items: PivotCacheSharedItems::default(),
        grouping: None,
    }
}

fn stats() -> PivotCacheSharedItemsStats {
    PivotCacheSharedItemsStats {
        text_field: true,
        non_dates: false,
        date_in_field: false,
        has_text_item: true,
        has_blank_item: true,
        mixed_types_ignoring_blanks: true,
        numeric_field: false,
        integer_field: false,
        has_long_text_item: false,
        item_count: 7,
        minimum: None,
        maximum: None,
    }
}

/// A field whose shared items cover every value type, both plain and
/// with additional info.
fn full_item_field() -> PivotCacheField {
    let mut region = field("Region");
    region.caption = Some("Sales Region".to_string());
    region.cant_get_unique_items = true;
    region.shared_items = PivotCacheSharedItems {
        stats: Some(stats()),
        items: vec![
            PivotCacheItem {
                value: PivotCacheItemValue::Missing,
                additional: None,
            },
            PivotCacheItem {
                value: PivotCacheItemValue::String("North".into()),
                additional: None,
            },
            PivotCacheItem {
                value: PivotCacheItemValue::String("South".into()),
                additional: Some(PivotCacheItemInfo {
                    ghost: true,
                    calculated: false,
                    caption: Some("SOUTH!".to_string()),
                    member_property_items: vec![1, -1],
                }),
            },
            PivotCacheItem {
                value: PivotCacheItemValue::Number(42.5),
                additional: None,
            },
            PivotCacheItem {
                value: PivotCacheItemValue::Boolean(true),
                additional: Some(PivotCacheItemInfo {
                    ghost: false,
                    calculated: true,
                    caption: None,
                    member_property_items: Vec::new(),
                }),
            },
            PivotCacheItem {
                value: PivotCacheItemValue::Error(PivotCacheErrorCode::NA),
                additional: None,
            },
            PivotCacheItem {
                value: PivotCacheItemValue::DateTime(PivotCacheDateTime {
                    year: 2024,
                    month: 3,
                    day: 14,
                    hour: 9,
                    minute: 30,
                    second: 0,
                }),
                additional: None,
            },
        ],
    };
    region
}

fn range_grouped_field() -> PivotCacheField {
    let mut amount = field("Amount");
    amount.number_format = Some(44);
    amount.grouping = Some(PivotCacheFieldGrouping {
        parent_field: None,
        base_field: Some(1),
        range: Some(PivotCacheRangeGrouping {
            group_by: PivotCacheGroupBy::Days,
            auto_start: true,
            auto_end: false,
            dates: true,
            start: 45_000.0,
            end: 46_000.0,
            interval: 7.0,
        }),
        discrete: None,
        items: Vec::new(),
    });
    amount
}

fn discrete_grouped_field() -> PivotCacheField {
    let mut group = field("RegionGroup");
    group.formula = Some(PivotParsedFormulaData {
        tokens: vec![0x1E, 0x02],
        extra: vec![0xAA],
    });
    group.member_property_fields = vec![2, 3];
    group.member_property_name = Some("Prop".to_string());
    group.ignore = true;
    group.grouping = Some(PivotCacheFieldGrouping {
        parent_field: Some(0),
        base_field: Some(0),
        range: None,
        discrete: Some(PivotCacheDiscreteGrouping {
            item_indexes: vec![1, 3, 5],
        }),
        items: vec![PivotCacheItem {
            value: PivotCacheItemValue::String("Grouped".into()),
            additional: None,
        }],
    });
    group
}

fn hierarchy() -> PivotCacheHierarchy {
    PivotCacheHierarchy {
        unique_name: "[Region]".to_string(),
        caption: "Region".to_string(),
        dimension_unique_name: Some("[RegionDim]".to_string()),
        default_member_unique_name: Some("[Region].[All]".to_string()),
        all_member_unique_name: None,
        all_member_display: Some("All Regions".to_string()),
        display_folder: None,
        measure_group: Some("MG".to_string()),
        measure: false,
        set: false,
        attribute_hierarchy: true,
        measure_hierarchy: false,
        only_one_field: true,
        time_hierarchy: false,
        key_attribute_hierarchy: true,
        hidden: false,
        unbalanced_real: Some(false),
        unbalanced_group: None,
        attribute_member_value_type: Some(0x0007),
        level_count: 2,
        set_parent_index: None,
        icon_set: -1,
        field_usage: vec![0, 1],
        grouping_levels: vec![PivotCacheGroupingLevel {
            group_level: true,
            unique_name: "[Region].[Custom]".to_string(),
            caption: "Custom".to_string(),
        }],
        grouping_groups: vec![PivotCacheGroupingGroup {
            group_number: 1,
            name: "Group1".to_string(),
            unique_name: "[Region].[Group1]".to_string(),
            caption: "Group 1".to_string(),
            parent_unique_name: Some("[Region].[All]".to_string()),
            members: vec![
                PivotCacheGroupingGroupMember {
                    is_group: false,
                    unique_name: "[Region].[North]".to_string(),
                },
                PivotCacheGroupingGroupMember {
                    is_group: true,
                    unique_name: "[Region].[South]".to_string(),
                },
            ],
        }],
        ext14: Some(PivotCacheHierarchyExt14 {
            flatten_hierarchies: true,
            measure_set: false,
            hierarchize_distinct: true,
            ignorable: false,
            hierarchy_indexes: vec![0, -2],
        }),
    }
}

fn full_definition() -> PivotCacheDefinition {
    PivotCacheDefinition {
        version_last_refresh: 3,
        version_refreshable_min: 0,
        version_created: 2,
        save_data: true,
        invalid: false,
        refresh_on_load: true,
        optimize_cache: false,
        enable_refresh: true,
        background_query: true,
        upgrade_on_refresh: false,
        cube_functions: true,
        support_subquery: true,
        support_attrib_drill: true,
        ghost_items_max: -1,
        refreshed_date_serial: 44_000.5,
        record_count: 5,
        refreshed_by: Some("analyst".to_string()),
        records_rel_id: Some("rIdRecords".to_string()),
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Worksheet,
            connection_id: None,
            worksheet: Some(PivotCacheWorksheetSource {
                named_range: None,
                built_in_name: false,
                sheet_name: Some("Data Sheet".to_string()),
                external_rel_id: None,
                range: Some(PivotCacheRange {
                    first_row: 0,
                    last_row: 99,
                    first_column: 1,
                    last_column: 7,
                }),
            }),
            consolidation: None,
        }),
        fields: vec![
            full_item_field(),
            range_grouped_field(),
            discrete_grouped_field(),
        ],
        hierarchies: vec![hierarchy()],
        tuple_cache: Some(PivotCacheTupleCache {
            entries: vec![
                PivotCacheItemValue::Missing,
                PivotCacheItemValue::Number(1.5),
                PivotCacheItemValue::String("cube".into()),
                PivotCacheItemValue::Boolean(false),
                PivotCacheItemValue::Error(PivotCacheErrorCode::Div0),
                PivotCacheItemValue::DateTime(PivotCacheDateTime {
                    year: 2023,
                    month: 12,
                    day: 31,
                    hour: 23,
                    minute: 59,
                    second: 59,
                }),
            ],
            queries: vec!["SELECT {} ON 0".to_string()],
            sets: vec![
                PivotCacheTupleCacheSet {
                    tuple_count: Some(4),
                    max_rank: 2,
                    sort_order: 1,
                    query_failed: false,
                    definition: "{[Region].Members}".to_string(),
                },
                PivotCacheTupleCacheSet {
                    tuple_count: None,
                    max_rank: 0,
                    sort_order: 0,
                    query_failed: true,
                    definition: "{}".to_string(),
                },
            ],
        }),
        calculated_items: vec![CalculatedItem {
            formula: PivotParsedFormulaData {
                tokens: vec![0x03, 0x04],
                extra: Vec::new(),
            },
            names: vec![PivotName {
                field_index: 0,
                function: PivotNameFunction::Sum,
                err_name: false,
                pairs: vec![PivotNamePair {
                    physical: true,
                    relative: false,
                    field_index: 0,
                    item_index: 2,
                }],
            }],
            filters: vec![PivotRuleFilter {
                field: -2,
                item_types: 0x1F,
                selected: true,
                items: vec![0, 2],
            }],
        }],
        calculated_members: vec![CalculatedMember {
            name: "[Measures].[Calc]".to_string(),
            mdx: "1+1".to_string(),
            solve_order: 5,
            is_set: true,
            member_name: Some("Calc".to_string()),
            source_hierarchy: Some("[Measures]".to_string()),
            parent_unique: Some("[Measures].[All]".to_string()),
            ext14: Some(CalculatedMemberExt14 {
                flatten_hierarchies: false,
                dynamic_set: true,
                hierarchize_distinct: true,
                display_folder: "Folder".to_string(),
                long_mdx: Some("1+1 /* long */".to_string()),
            }),
        }],
        ext14: Some(PivotCacheDefinitionExt14 {
            slicer_data: true,
            server_support_subquery_calc_mem: true,
            server_support_subquery_non_visual: false,
            server_support_add_calc_mems: true,
            cache_id: 12,
        }),
    }
}

fn round_trip(definition: &PivotCacheDefinition) -> PivotCacheDefinition {
    let bytes = write_pivot_cache_definition(definition).unwrap();
    parse_pivot_cache_definition(&bytes).unwrap()
}

#[test]
fn serialized_full_definition_round_trips_through_the_reader() {
    let definition = full_definition();
    assert_eq!(round_trip(&definition), definition);
}

#[test]
fn serialized_minimal_definition_round_trips() {
    let definition = PivotCacheDefinition::default();
    assert_eq!(round_trip(&definition), definition);
}

#[test]
fn serialized_consolidation_source_round_trips() {
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Consolidation,
            connection_id: None,
            worksheet: None,
            consolidation: Some(PivotCacheConsolidationSource {
                auto_page: true,
                sets: vec![
                    PivotCacheConsolidationSet {
                        item_indexes: [1, u32::MAX, u32::MAX, u32::MAX],
                        named_range: Some("MyRange".to_string()),
                        built_in_name: true,
                        sheet_name: None,
                        external_rel_id: None,
                        range: None,
                    },
                    PivotCacheConsolidationSet {
                        item_indexes: [0, 1, u32::MAX, u32::MAX],
                        named_range: None,
                        built_in_name: false,
                        sheet_name: Some("Q1".to_string()),
                        external_rel_id: Some("rIdExt".to_string()),
                        range: Some(PivotCacheRange {
                            first_row: 4,
                            last_row: 20,
                            first_column: 0,
                            last_column: 3,
                        }),
                    },
                ],
                pages: vec![
                    PivotCacheConsolidationPage {
                        item_names: vec!["Region1".to_string(), "Region2".to_string()],
                    },
                    PivotCacheConsolidationPage {
                        item_names: Vec::new(),
                    },
                ],
            }),
        }),
        ..PivotCacheDefinition::default()
    };
    assert_eq!(round_trip(&definition), definition);
}

#[test]
fn serialized_external_source_and_named_range_round_trips() {
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::External,
            connection_id: Some(17),
            worksheet: Some(PivotCacheWorksheetSource {
                named_range: Some("ExternalData".to_string()),
                built_in_name: false,
                sheet_name: None,
                external_rel_id: Some("rIdBook".to_string()),
                range: None,
            }),
            consolidation: None,
        }),
        ..PivotCacheDefinition::default()
    };
    assert_eq!(round_trip(&definition), definition);
}

#[test]
fn refuses_content_that_cannot_round_trip() {
    // Index items inside shared items are skipped by the reader.
    let mut definition = PivotCacheDefinition::default();
    let mut broken = field("F");
    broken.shared_items = PivotCacheSharedItems {
        stats: Some(stats()),
        items: vec![PivotCacheItem {
            value: PivotCacheItemValue::Index(3),
            additional: None,
        }],
    };
    definition.fields = vec![broken];
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Shared items without statistics would fabricate an ATBL payload.
    let mut definition = PivotCacheDefinition::default();
    let mut broken = field("F");
    broken.shared_items = PivotCacheSharedItems {
        stats: None,
        items: vec![PivotCacheItem {
            value: PivotCacheItemValue::Number(1.0),
            additional: None,
        }],
    };
    definition.fields = vec![broken];
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Statistics with only one bound set.
    let mut definition = PivotCacheDefinition::default();
    let mut broken = field("F");
    broken.shared_items = PivotCacheSharedItems {
        stats: Some(PivotCacheSharedItemsStats {
            minimum: Some(1.0),
            ..stats()
        }),
        items: Vec::new(),
    };
    definition.fields = vec![broken];
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Index items inside grouping items.
    let mut definition = PivotCacheDefinition::default();
    let mut broken = field("F");
    broken.grouping = Some(PivotCacheFieldGrouping {
        parent_field: None,
        base_field: None,
        range: None,
        discrete: None,
        items: vec![PivotCacheItem {
            value: PivotCacheItemValue::Index(0),
            additional: None,
        }],
    });
    definition.fields = vec![broken];
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Index values inside tuple cache entries.
    let definition = PivotCacheDefinition {
        tuple_cache: Some(PivotCacheTupleCache {
            entries: vec![PivotCacheItemValue::Index(1)],
            ..PivotCacheTupleCache::default()
        }),
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn refuses_ambiguous_or_empty_sources() {
    // Both a named range and a cell range.
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Worksheet,
            connection_id: None,
            worksheet: Some(PivotCacheWorksheetSource {
                named_range: Some("R".to_string()),
                built_in_name: false,
                sheet_name: None,
                external_rel_id: None,
                range: Some(PivotCacheRange {
                    first_row: 0,
                    last_row: 1,
                    first_column: 0,
                    last_column: 0,
                }),
            }),
            consolidation: None,
        }),
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Neither a named range nor a cell range.
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Worksheet,
            connection_id: None,
            worksheet: Some(PivotCacheWorksheetSource {
                named_range: None,
                built_in_name: false,
                sheet_name: None,
                external_rel_id: None,
                range: None,
            }),
            consolidation: None,
        }),
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // A consolidation set with no locator.
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Consolidation,
            connection_id: None,
            worksheet: None,
            consolidation: Some(PivotCacheConsolidationSource {
                auto_page: false,
                sets: vec![PivotCacheConsolidationSet {
                    item_indexes: [u32::MAX; 4],
                    named_range: None,
                    built_in_name: false,
                    sheet_name: None,
                    external_rel_id: None,
                    range: None,
                }],
                pages: Vec::new(),
            }),
        }),
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn refuses_out_of_range_model_values() {
    // More than four consolidation pages.
    let definition = PivotCacheDefinition {
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Consolidation,
            connection_id: None,
            worksheet: None,
            consolidation: Some(PivotCacheConsolidationSource {
                auto_page: false,
                sets: Vec::new(),
                pages: vec![
                    PivotCacheConsolidationPage {
                        item_names: Vec::new(),
                    };
                    5
                ],
            }),
        }),
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));

    // Rule-filter item types exceeding the 13-bit mask.
    let definition = PivotCacheDefinition {
        calculated_items: vec![CalculatedItem {
            formula: PivotParsedFormulaData::default(),
            names: Vec::new(),
            filters: vec![PivotRuleFilter {
                field: 0,
                item_types: 1 << 13,
                selected: false,
                items: Vec::new(),
            }],
        }],
        ..PivotCacheDefinition::default()
    };
    assert!(matches!(
        write_pivot_cache_definition(&definition),
        Err(Error::Unrecognized { .. })
    ));
}

#[test]
fn written_caches_round_trip_through_the_package() {
    let first = full_definition();
    let second = PivotCacheDefinition {
        record_count: 2,
        source: Some(PivotCacheSource {
            source_type: PivotCacheSourceType::Worksheet,
            connection_id: None,
            worksheet: Some(PivotCacheWorksheetSource {
                named_range: Some("Data".to_string()),
                built_in_name: false,
                sheet_name: Some("Sheet1".to_string()),
                external_rel_id: None,
                range: None,
            }),
            consolidation: None,
        }),
        ..PivotCacheDefinition::default()
    };

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
    let first_id = workbook.add_pivot_cache(&first).unwrap();
    let second_id = workbook.add_pivot_cache(&second).unwrap();
    assert_eq!(first_id, 1);
    assert_eq!(second_id, 2);

    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    let reader = crate::package::Workbook::new(Cursor::new(output.into_inner())).unwrap();

    let definitions = reader.pivot_cache_definitions();
    assert_eq!(definitions.len(), 2);
    assert_eq!(definitions[0].0, first_id);
    assert_eq!(definitions[1].0, second_id);
    assert_eq!(definitions[0].1, first);
    assert_eq!(definitions[1].1, second);
    assert_eq!(reader.pivot_cache_definition(first_id), Some(&first));
    assert!(reader.pivot_cache_definition(99).is_none());
}

#[test]
fn add_pivot_cache_surfaces_serializer_refusals() {
    let mut broken = PivotCacheDefinition::default();
    let mut broken_field = field("F");
    broken_field.shared_items = PivotCacheSharedItems {
        stats: None,
        items: vec![PivotCacheItem {
            value: PivotCacheItemValue::Number(1.0),
            additional: None,
        }],
    };
    broken.fields = vec![broken_field];

    let mut workbook = WorkbookWriter::new();
    workbook.add_worksheet(MutableWorksheet::new("Sheet1"));
    assert!(matches!(
        workbook.add_pivot_cache(&broken),
        Err(Error::Unrecognized { .. })
    ));
    // The refused cache was not attached; a valid cache gets id 1.
    let valid = PivotCacheDefinition::default();
    assert_eq!(workbook.add_pivot_cache(&valid).unwrap(), 1);
}
