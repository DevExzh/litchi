use super::model::{compare_datetime, validate_datetime};
use super::*;
use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI};

#[cfg(test)]
#[allow(
    clippy::module_inception,
    reason = "the nested module keeps timeline regressions scoped under the test support module"
)]
mod tests {
    use super::*;

    fn state() -> State {
        State {
            selection: Some(Range::new("2024-01-02T00:00:00Z", "2024-01-31T23:59:59Z").unwrap()),
            bounds: Range::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z").unwrap(),
            extension_list: None,
            single_range_filter_state: None,
            minimal_refresh_version: 0,
            last_refresh_version: 1,
            pivot_cache_id: 1,
            filter_type: FilterType::DateBetween,
        }
    }
    fn cache() -> Cache {
        Cache {
            relationship_id: "rIdViewCache1".into(),
            part_name: "/xl/timelineCaches/timelineCache1.xml".into(),
            definition: CacheDefinition {
                name: "View_Date".into(),
                uid: Some("{11111111-1111-1111-1111-111111111111}".into()),
                source_name: "Date".into(),
                pivot_tables: vec![CachePivotTable {
                    tab_id: 1,
                    name: "PivotTable1".into(),
                }],
                state: state(),
                timeline_pivot_filter: None,
                extension_list: None,
            },
        }
    }
    fn views() -> Views {
        Views {
            timelines: vec![View {
                name: "View_Date_View".into(),
                uid: Some("{22222222-2222-2222-2222-222222222222}".into()),
                cache: "View_Date".into(),
                caption: Some("Date".into()),
                show_header: Some(true),
                show_selection_label: Some(false),
                show_time_level: None,
                show_horizontal_scrollbar: Some(true),
                level: Level::Month,
                selection_level: Level::Day,
                scroll_position: Some("2024-01-02T03:04:05Z".into()),
                style: Some("ViewStyleLight1".into()),
                extension_list: None,
            }],
        }
    }
    fn package() -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let workbook = PackURI::new("/xl/workbook.xml").unwrap();
        let sheet = PackURI::new("/xl/worksheets/sheet1.xml").unwrap();
        package.add_part(Box::new(BlobPart::new(
            workbook.clone(),
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml".into(),
            format!(r#"<workbook xmlns="{SML}"><sheets/></workbook>"#).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            sheet,
            ct::SML_WORKSHEET.into(),
            format!(r#"<worksheet xmlns="{SML}"><sheetData/></worksheet>"#).into_bytes(),
        )));
        (package, workbook)
    }

    #[test]
    fn typed_protocol_shaped_parts_round_trip() {
        let cache = cache().definition;
        let cache_xml = write_timeline_cache_definition(&cache).unwrap();
        assert_eq!(parse_timeline_cache_definition(&cache_xml).unwrap(), cache);
        let timelines = views();
        let xml = write_timelines(&timelines).unwrap();
        assert_eq!(parse_timelines(&xml).unwrap(), timelines);
    }

    #[test]
    fn compatibility_directives_keep_known_revision_ids_and_reject_unknown_requirements() {
        let value = cache().definition;
        let xml = String::from_utf8(write_timeline_cache_definition(&value).unwrap()).unwrap();
        let with_mce = xml.replace(
            "<x15:timelineCacheDefinition ",
            r#"<x15:timelineCacheDefinition xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="xr10" "#,
        );
        assert_eq!(
            parse_timeline_cache_definition(with_mce.as_bytes()).unwrap(),
            value
        );
        let required = with_mce.replace(
            "mc:Ignorable=\"xr10\"",
            r#"xmlns:future="urn:future" mc:MustUnderstand="future""#,
        );
        assert!(parse_timeline_cache_definition(required.as_bytes()).is_err());
        let undeclared = with_mce.replace("mc:Ignorable=\"xr10\"", "mc:Ignorable=\"missing\"");
        assert!(parse_timeline_cache_definition(undeclared.as_bytes()).is_err());
    }

    #[test]
    fn documented_2010_timeline_graph_aliases_remain_readable_and_owned() {
        let (mut package, workbook) = package();
        let cache = cache();
        store_timeline_caches(&mut package, &workbook, std::slice::from_ref(&cache)).unwrap();
        let view = Part {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdView1".into(),
            part_name: "/xl/timelines/timeline1.xml".into(),
            timelines: views(),
        };
        store_worksheet_timelines(&mut package, &workbook, &view).unwrap();
        for (owner, id, target, content_type, relationship_type) in [
            (
                "/xl/workbook.xml",
                "rIdViewCache1",
                "/xl/timelineCaches/timelineCache1.xml",
                "application/vnd.ms-excel.TimelineCache+xml",
                "http://schemas.microsoft.com/office/2010/relationships/TimelineCache",
            ),
            (
                "/xl/worksheets/sheet1.xml",
                "rIdView1",
                "/xl/timelines/timeline1.xml",
                "application/vnd.ms-excel.Timeline+xml",
                "http://schemas.microsoft.com/office/2010/relationships/Timeline",
            ),
        ] {
            let target = PackURI::new(target).unwrap();
            let owner = PackURI::new(owner).unwrap();
            let bytes = package.get_part(&target).unwrap().blob().to_vec();
            assert!(package.remove_part(&target));
            package
                .try_add_part(Box::new(BlobPart::new(
                    target.clone(),
                    content_type.into(),
                    bytes,
                )))
                .unwrap();
            let rels = package.get_part_mut(&owner).unwrap().rels_mut();
            rels.remove(id).unwrap();
            rels.try_add_relationship(
                relationship_type.into(),
                target.relative_ref(owner.base_uri()),
                id.into(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
        }
        assert_eq!(load_timeline_caches(&package, &workbook).unwrap(), [cache]);
        assert_eq!(load_timelines(&package, &workbook).unwrap(), [view]);
        package
            .rels_mut()
            .try_add_relationship(
                "http://schemas.microsoft.com/office/2010/relationships/TimelineCache".into(),
                "xl/timelineCaches/timelineCache1.xml".into(),
                "orphanAlias".into(),
                litchi_opc::TargetMode::Internal,
            )
            .unwrap();
        assert!(load_timeline_caches(&package, &workbook).is_err());
    }

    #[test]
    fn typed_state_pivot_filter_and_all_filter_types_round_trip() {
        let tokens = [
            "unknown",
            "count",
            "percent",
            "sum",
            "captionEqual",
            "captionNotEqual",
            "captionBeginsWith",
            "captionNotBeginsWith",
            "captionEndsWith",
            "captionNotEndsWith",
            "captionContains",
            "captionNotContains",
            "captionGreaterThan",
            "captionGreaterThanOrEqual",
            "captionLessThan",
            "captionLessThanOrEqual",
            "captionBetween",
            "captionNotBetween",
            "valueEqual",
            "valueNotEqual",
            "valueGreaterThan",
            "valueGreaterThanOrEqual",
            "valueLessThan",
            "valueLessThanOrEqual",
            "valueBetween",
            "valueNotBetween",
            "dateEqual",
            "dateNotEqual",
            "dateOlderThan",
            "dateOlderThanOrEqual",
            "dateNewerThan",
            "dateNewerThanOrEqual",
            "dateBetween",
            "dateNotBetween",
            "tomorrow",
            "today",
            "yesterday",
            "nextWeek",
            "thisWeek",
            "lastWeek",
            "nextMonth",
            "thisMonth",
            "lastMonth",
            "nextQuarter",
            "thisQuarter",
            "lastQuarter",
            "nextYear",
            "thisYear",
            "lastYear",
            "yearToDate",
            "Q1",
            "Q2",
            "Q3",
            "Q4",
            "M1",
            "M2",
            "M3",
            "M4",
            "M5",
            "M6",
            "M7",
            "M8",
            "M9",
            "M10",
            "M11",
            "M12",
        ];
        for token in tokens {
            assert_eq!(FilterType::parse(token).unwrap().as_str(), token)
        }
        let xml = format!(
            r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" xmlns:x="{SML}" name="View_Date" sourceName="Date"><x15:state minimalRefreshVersion="1" lastRefreshVersion="2" pivotCacheId="3" filterType="count"><x15:selection startDate="2024-02-01T00:00:00+02:00" endDate="2024-02-29T23:59:59+02:00"/><x15:bounds startDate="2024-01-01T00:00:00+02:00" endDate="2024-12-31T23:59:59+02:00"/></x15:state><x15:timelinePivotFilter useWholeDay="1" fld="2" id="7" name="Recent"><x:autoFilter ref="A1:B9"><x:filterColumn colId="0"><x:customFilters><x:customFilter operator="greaterThan" val="2"/></x:customFilters></x:filterColumn></x:autoFilter></x15:timelinePivotFilter></x15:timelineCacheDefinition>"#
        );
        let parsed = parse_timeline_cache_definition(xml.as_bytes()).unwrap();
        assert!(
            parsed
                .timeline_pivot_filter
                .as_ref()
                .unwrap()
                .auto_filter
                .is_some()
        );
        let written = write_timeline_cache_definition(&parsed).unwrap();
        assert_eq!(parse_timeline_cache_definition(&written).unwrap(), parsed);
    }

    #[test]
    fn rejects_state_calendar_timezone_range_order_and_filter_presence() {
        let state = |filter: &str, selection: &str, bounds: &str, pivot: &str| {
            format!(
                r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" name="View_Date" sourceName="Date"><x15:state minimalRefreshVersion="0" lastRefreshVersion="0" pivotCacheId="1" filterType="{filter}">{selection}{bounds}</x15:state>{pivot}</x15:timelineCacheDefinition>"#
            )
        };
        let good_bounds =
            r#"<x15:bounds startDate="2024-01-01T00:00:00Z" endDate="2024-12-31T23:59:59Z"/>"#;
        for xml in [
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2024-02-30T00:00:00Z" endDate="2024-12-31T00:00:00Z"/>"#,
                "",
            ),
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2024-01-01T00:00:00+15:00" endDate="2024-12-31T00:00:00+15:00"/>"#,
                "",
            ),
            state(
                "dateBetween",
                "",
                r#"<x15:bounds startDate="2025-01-01T00:00:00Z" endDate="2024-01-01T00:00:00Z"/>"#,
                "",
            ),
            state(
                "dateBetween",
                r#"<x15:selection startDate="2023-01-01T00:00:00Z" endDate="2024-02-01T00:00:00Z"/>"#,
                good_bounds,
                "",
            ),
            state(
                "dateBetween",
                "",
                good_bounds,
                r#"<x15:timelinePivotFilter fld="0" id="1"/>"#,
            ),
            state("bogus", "", good_bounds, ""),
        ] {
            assert!(
                parse_timeline_cache_definition(xml.as_bytes()).is_err(),
                "accepted {xml}"
            );
        }
    }

    #[test]
    fn package_store_load_and_cross_validation_round_trip() {
        let (mut package, workbook) = package();
        let expected_cache = cache();
        store_timeline_caches(
            &mut package,
            &workbook,
            std::slice::from_ref(&expected_cache),
        )
        .unwrap();
        assert_eq!(
            load_timeline_caches(&package, &workbook).unwrap(),
            vec![expected_cache]
        );
        let expected = Part {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdView1".into(),
            part_name: "/xl/timelines/timeline1.xml".into(),
            timelines: views(),
        };
        store_worksheet_timelines(&mut package, &workbook, &expected).unwrap();
        assert_eq!(load_timelines(&package, &workbook).unwrap(), vec![expected]);
    }

    #[test]
    fn rejects_hostile_grammar_identity_and_bounds() {
        for xml in [
            format!(r#"<!DOCTYPE x><x15:timelines xmlns:x15="{X15}"/>"#),
            format!(
                r#"<x15:timelines xmlns:x15="{X15}"><x15:timeline name="x" cache="c" level="4" selectionLevel="0"/></x15:timelines>"#
            ),
            format!(
                r#"<x15:timelineCacheDefinition xmlns:x15="{X15}" name="A1" sourceName="Date"><x15:state/></x15:timelineCacheDefinition>"#
            ),
        ] {
            assert!(
                parse_timelines(xml.as_bytes()).is_err()
                    || parse_timeline_cache_definition(xml.as_bytes()).is_err()
            );
        }
        assert!(parse_timelines(&vec![b' '; MAX_XML_BYTES + 1]).is_err());
        let mut duplicate = views();
        duplicate.timelines.push(duplicate.timelines[0].clone());
        assert!(write_timelines(&duplicate).is_err());
    }

    #[test]
    fn rejects_package_graph_and_unknown_cache_errors() {
        let (mut package, workbook) = package();
        store_timeline_caches(&mut package, &workbook, &[cache()]).unwrap();
        let mut bad = Part {
            worksheet_part_name: "/xl/worksheets/sheet1.xml".into(),
            relationship_id: "rIdView1".into(),
            part_name: "/xl/timelines/timeline1.xml".into(),
            timelines: views(),
        };
        bad.timelines.timelines[0].cache = "Missing".into();
        assert!(store_worksheet_timelines(&mut package, &workbook, &bad).is_err());
        let target = PackURI::new("/xl/timelineCaches/timelineCache1.xml").unwrap();
        package
            .get_part_mut(&target)
            .unwrap()
            .rels_mut()
            .add_relationship(
                "urn:forbidden".into(),
                "x.xml".into(),
                "rIdBad".into(),
                false,
            );
        assert!(load_timeline_caches(&package, &workbook).is_err());
    }
}

#[cfg(test)]
mod xsd_datetime_full_lexical_tests {
    use super::*;

    #[test]
    fn accepts_and_compares_full_applicable_xsd_datetime_space() {
        for value in [
            "2000-02-29T24:00:00Z",
            "-0001-02-29T24:00:00-14:00",
            "12345-12-31T23:59:59.12345678901234567890+14:00",
            "-12345-01-01T00:00:00.00000000000000000001",
        ] {
            validate_datetime(value).unwrap();
        }
        assert_eq!(
            compare_datetime("2000-02-29T24:00:00Z", "2000-03-01T00:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("0001-01-01T00:00:00+14:00", "-0001-12-31T10:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("9999-12-31T24:00:00+14:00", "9999-12-31T10:00:00Z").unwrap(),
            std::cmp::Ordering::Equal
        );
        assert_eq!(
            compare_datetime("-10000-01-01T00:00:00Z", "-9999-01-01T00:00:00Z").unwrap(),
            std::cmp::Ordering::Less
        );
    }

    #[test]
    fn rejects_non_xsd_datetime_lexemes_and_indeterminate_mixed_zones() {
        for value in [
            "0000-01-01T00:00:00Z",
            "-0000-01-01T00:00:00Z",
            "999-01-01T00:00:00Z",
            "01234-01-01T00:00:00Z",
            "0001-02-29T00:00:00Z",
            "-0002-02-29T00:00:00Z",
            "2000-01-01T24:00:00.1Z",
            "2000-01-01T24:00:01Z",
            "2000-01-01t00:00:00Z",
            "2000-01-01T00:00:00z",
            "2000-01-01T00:00:60Z",
            "2000-01-01T00:00:00+14:01",
            "2000-01-01T00:00:00-15:00",
        ] {
            assert!(validate_datetime(value).is_err(), "accepted {value}");
        }
        assert!(compare_datetime("2000-01-01T00:00:00", "2000-01-01T00:00:00Z").is_err());
    }

    #[test]
    fn timeline_range_preserves_extended_lexemes_round_trip() {
        let range = Range::new("-12345-01-01T24:00:00", "12345-12-31T24:00:00").unwrap();
        assert_eq!(range.start_date(), "-12345-01-01T24:00:00");
        assert_eq!(range.end_date(), "12345-12-31T24:00:00");
    }
}

#[cfg(test)]
mod child_order_tests {
    use super::*;

    /// Every sequence of at most `max_len` child indices below `count`,
    /// shortest first.
    fn sequences(count: usize, max_len: usize) -> Vec<Vec<usize>> {
        let mut all = vec![Vec::new()];
        let mut start = 0;
        for _ in 0..max_len {
            let end = all.len();
            for index in start..end {
                for child in 0..count {
                    let mut next = all[index].clone();
                    next.push(child);
                    all.push(next);
                }
            }
            start = end;
        }
        all
    }

    fn assert_rows(actual: &[(String, String)], expected: &[(&str, &str)]) {
        let same = actual.len() == expected.len()
            && actual.iter().zip(expected).all(
                |((label, outcome), (expected_label, expected_outcome))| {
                    label == expected_label && outcome == expected_outcome
                },
            );
        let observed: String = actual
            .iter()
            .map(|(label, outcome)| format!("({label:?}, {outcome:?}),\n"))
            .collect();
        assert!(same, "observed rows:\n{observed}");
    }

    fn outcome<T>(result: Result<T>) -> String {
        match result {
            Ok(_) => "ok".to_owned(),
            Err(error) => format!("{error:?}"),
        }
    }

    const BOUNDS: &str =
        "<x15:bounds startDate=\"2024-01-01T00:00:00Z\" endDate=\"2024-12-31T23:59:59Z\"/>";

    #[test]
    fn state_child_order_matrix_is_stable() {
        let children = [
            (
                "selection",
                "<x15:selection startDate=\"2024-01-02T00:00:00Z\" endDate=\"2024-01-31T23:59:59Z\"/>",
            ),
            ("bounds", BOUNDS),
            ("extLst", "<x15:extLst/>"),
        ];
        let mut cases: Vec<Vec<(&str, &str)>> = sequences(children.len(), 3)
            .into_iter()
            .map(|sequence| sequence.into_iter().map(|index| children[index]).collect())
            .collect();
        cases.push(vec![("unknown", "<x15:unknown/>")]);
        cases.push(vec![children[1], ("unknown", "<x15:unknown/>")]);
        cases.push(vec![children[1], ("core extLst", "<x:extLst/>")]);
        cases.push(vec![("core bounds", "<x:bounds/>")]);
        let actual: Vec<(String, String)> = cases
            .iter()
            .map(|case| {
                let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
                let body: String = case.iter().map(|(_, xml)| *xml).collect();
                let xml = format!(
                    "<x15:timelineCacheDefinition xmlns:x15=\"{X15}\" xmlns:x=\"{SML}\" name=\"View_Date\" sourceName=\"Date\"><x15:state minimalRefreshVersion=\"6\" lastRefreshVersion=\"6\" pivotCacheId=\"1\" filterType=\"dateBetween\">{body}</x15:state></x15:timelineCacheDefinition>"
                );
                (
                    label.join(","),
                    outcome(parse_timeline_cache_definition(xml.as_bytes())),
                )
            })
            .collect();
        assert_rows(
            &actual,
            &[
                ("", "Invalid(\"View state requires bounds\")"),
                ("selection", "Invalid(\"View state requires bounds\")"),
                ("bounds", "ok"),
                (
                    "extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                ("selection,bounds", "ok"),
                (
                    "selection,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                ("bounds,extLst", "ok"),
                (
                    "extLst,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,selection,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,selection,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,selection,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,bounds,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,bounds,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                ("selection,bounds,extLst", "ok"),
                (
                    "selection,extLst,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,extLst,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "selection,extLst,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,selection,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,selection,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,selection,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,bounds,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,bounds,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,bounds,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,extLst,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,extLst,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,extLst,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,selection,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,selection,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,selection,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,bounds,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,bounds,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,bounds,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,extLst,selection",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,extLst,bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "extLst,extLst,extLst",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "unknown",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                (
                    "bounds,unknown",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
                ("bounds,core extLst", "ok"),
                (
                    "core bounds",
                    "Invalid(\"invalid or out-of-order View state child\")",
                ),
            ],
        );
    }

    #[test]
    fn cache_definition_child_order_matrix_is_stable() {
        let state = format!(
            "<x15:state minimalRefreshVersion=\"6\" lastRefreshVersion=\"6\" pivotCacheId=\"1\" filterType=\"thisMonth\">{BOUNDS}</x15:state>"
        );
        let children = [
            (
                "pivotTables",
                "<x15:pivotTables><x15:pivotTable tabId=\"1\" name=\"PivotTable1\"/></x15:pivotTables>",
            ),
            ("state", state.as_str()),
            ("filter", "<x15:timelinePivotFilter fld=\"0\" id=\"1\"/>"),
            ("extLst", "<x15:extLst/>"),
        ];
        let mut cases: Vec<Vec<(&str, &str)>> = sequences(children.len(), 3)
            .into_iter()
            .map(|sequence| sequence.into_iter().map(|index| children[index]).collect())
            .collect();
        // Every pair again behind the required state, so that an accepted
        // order is not hidden by the missing-state refusal.
        for sequence in sequences(children.len(), 2)
            .into_iter()
            .filter(|sequence| sequence.len() == 2)
        {
            let mut case = vec![children[0], children[1]];
            case.extend(sequence.into_iter().map(|index| children[index]));
            cases.push(case);
        }
        cases.push(vec![("unknown", "<x15:unknown/>")]);
        cases.push(vec![children[1], ("unknown", "<x15:unknown/>")]);
        cases.push(vec![children[1], ("core extLst", "<x:extLst/>")]);
        cases.push(vec![("core state", "<x:state/>")]);
        let actual: Vec<(String, String)> = cases
            .iter()
            .map(|case| {
                let label = case.iter().map(|(name, _)| *name).collect::<Vec<_>>();
                let body: String = case.iter().map(|(_, xml)| *xml).collect();
                let xml = format!(
                    "<x15:timelineCacheDefinition xmlns:x15=\"{X15}\" xmlns:x=\"{SML}\" name=\"View_Date\" sourceName=\"Date\">{body}</x15:timelineCacheDefinition>"
                );
                (
                    label.join(","),
                    outcome(parse_timeline_cache_definition(xml.as_bytes())),
                )
            })
            .collect();
        assert_rows(
            &actual,
            &[
                (
                    "",
                    "Invalid(\"View Cache requires exactly one state element\")",
                ),
                (
                    "pivotTables",
                    "Invalid(\"View Cache requires exactly one state element\")",
                ),
                ("state", "ok"),
                (
                    "filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                ("pivotTables,state", "ok"),
                (
                    "pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                ("state,filter", "ok"),
                ("state,extLst", "ok"),
                (
                    "filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,pivotTables,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                ("pivotTables,state,filter", "ok"),
                ("pivotTables,state,extLst", "ok"),
                (
                    "pivotTables,filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,filter,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "state,pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,pivotTables,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,state,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,state,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                ("state,filter,extLst", "ok"),
                (
                    "state,extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "state,extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "state,extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "state,extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "filter,pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,pivotTables,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,state,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,state,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,filter,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "filter,extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "extLst,pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,pivotTables,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,state,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,state,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,filter,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "extLst,extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "pivotTables,state,pivotTables,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,pivotTables,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,pivotTables,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,pivotTables,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,state,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,state,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,state,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,state,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,filter,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,filter,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,filter,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                ("pivotTables,state,filter,extLst", "ok"),
                (
                    "pivotTables,state,extLst,pivotTables",
                    "Invalid(\"unexpected or out-of-order View Cache child 'pivotTables'\")",
                ),
                (
                    "pivotTables,state,extLst,state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
                (
                    "pivotTables,state,extLst,filter",
                    "Invalid(\"unexpected or out-of-order View Cache child 'timelinePivotFilter'\")",
                ),
                (
                    "pivotTables,state,extLst,extLst",
                    "Invalid(\"unexpected or out-of-order View Cache child 'extLst'\")",
                ),
                (
                    "unknown",
                    "Invalid(\"unexpected or out-of-order View Cache child 'unknown'\")",
                ),
                (
                    "state,unknown",
                    "Invalid(\"unexpected or out-of-order View Cache child 'unknown'\")",
                ),
                ("state,core extLst", "ok"),
                (
                    "core state",
                    "Invalid(\"unexpected or out-of-order View Cache child 'state'\")",
                ),
            ],
        );
    }
}
