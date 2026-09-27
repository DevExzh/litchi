use quick_xml::events::BytesStart;

use super::{
    AttributeCensusTermination0798, BytesStartExt, begin_attribute_census_0798,
    finish_attribute_census_0798,
};

fn tag(content: &str) -> BytesStart<'static> {
    let name_len = content
        .find([' ', '\t', '\r', '\n'])
        .unwrap_or(content.len());
    BytesStart::from_content(content.to_owned(), name_len)
}

fn only_row(report: super::AttributeCensusReport0798) -> super::AttributeCensusRow0798 {
    assert_eq!(report.rows.len(), 1, "{report:?}");
    assert_eq!(report.live_instances_at_finish, 0, "{report:?}");
    report.rows.into_iter().next().unwrap()
}

#[test]
fn disabled_census_does_not_retain_a_row() {
    let tag = tag("e a=\"1\"");
    let mut attributes = tag.checked_attributes();
    assert!(attributes.next().is_some());
    drop(attributes);
    let report = finish_attribute_census_0798();
    assert_eq!(report, super::AttributeCensusReport0798::default());
}

#[test]
fn full_consumption_records_end_and_unchecked_lexical_scan() {
    begin_attribute_census_0798();
    let tag = tag("e a=\"1\" b=\"2\"");
    let mut attributes = tag.checked_attributes();
    assert!(attributes.next().unwrap().is_ok());
    assert!(attributes.next().unwrap().is_ok());
    assert!(attributes.next().is_none());
    drop(attributes);
    let row = only_row(finish_attribute_census_0798());
    assert_eq!(row.element_name, b"e");
    assert_eq!(row.source, b"e a=\"1\" b=\"2\"");
    assert_eq!(row.next_calls, 3);
    assert_eq!(row.successful_yields, 2);
    assert_eq!(row.error_yields, 0);
    assert_eq!(row.end_yields, 1);
    assert_eq!(row.lexical_attribute_count, 2);
    assert_eq!(row.lexical_item_count, 2);
    assert_eq!(row.lexical_error_count, 0);
    assert!(row.lexical_scan_completed);
    assert_eq!(row.termination, AttributeCensusTermination0798::Exhausted);
    assert!(row.dropped);
    assert!(!row.early_drop);
    assert!(!row.partial_consumption);
    assert!(!row.counter_saturated);
}

#[test]
fn first_only_and_no_consumption_are_partial_early_drops() {
    begin_attribute_census_0798();
    let tag = tag("e a=\"1\" b=\"2\"");
    let mut attributes = tag.checked_attributes();
    assert!(attributes.next().unwrap().is_ok());
    drop(attributes);
    let row = only_row(finish_attribute_census_0798());
    assert_eq!(row.successful_yields, 1);
    assert!(row.early_drop);
    assert!(row.partial_consumption);
    assert_eq!(row.termination, AttributeCensusTermination0798::Dropped);

    begin_attribute_census_0798();
    let tag = tag("e a=\"1\"");
    let attributes = tag.checked_attributes();
    drop(attributes);
    let row = only_row(finish_attribute_census_0798());
    assert_eq!(row.next_calls, 0);
    assert!(row.early_drop);
    assert!(row.partial_consumption);
}

#[test]
fn checked_error_is_terminal_even_when_unchecked_scan_finds_tail() {
    begin_attribute_census_0798();
    let tag = tag("e a=\"1\" a=\"2\" tail=\"ok\"");
    let mut attributes = tag.checked_attributes();
    assert!(attributes.next().unwrap().is_ok());
    assert!(attributes.next().unwrap().is_err());
    drop(attributes);
    let row = only_row(finish_attribute_census_0798());
    assert_eq!(row.successful_yields, 1);
    assert_eq!(row.error_yields, 1);
    assert_eq!(row.end_yields, 0);
    assert_eq!(row.lexical_attribute_count, 3);
    assert_eq!(row.lexical_item_count, 3);
    assert_eq!(row.termination, AttributeCensusTermination0798::Error);
    assert!(!row.early_drop);
    assert!(row.partial_consumption);
}

#[test]
fn clone_resets_local_counts_and_preserves_prefix_lineage() {
    begin_attribute_census_0798();
    let tag = tag("e a=\"1\" b=\"2\"");
    let mut original = tag.checked_attributes();
    assert!(original.next().unwrap().is_ok());
    let mut clone = original.clone();
    assert!(clone.next().unwrap().is_ok());
    assert!(clone.next().is_none());
    drop(clone);
    drop(original);
    let report = finish_attribute_census_0798();
    assert_eq!(report.iterator_starts, 1);
    assert_eq!(report.iterator_clones, 1);
    assert_eq!(report.iterator_drops, 2);
    assert_eq!(report.live_instances_at_finish, 0);
    assert_eq!(report.rows.len(), 2);
    let original = &report.rows[0];
    let clone = &report.rows[1];
    assert_eq!(original.lineage_id, clone.lineage_id);
    assert!(!original.is_clone);
    assert!(clone.is_clone);
    assert_eq!(clone.starting_successful_yields, 1);
    assert_eq!(clone.successful_yields, 1);
    assert_eq!(clone.next_calls, 2);
    assert!(!clone.partial_consumption);
}

#[test]
fn stale_iterators_from_a_previous_session_are_not_counted() {
    begin_attribute_census_0798();
    let stale_tag = tag("e stale=\"1\"");
    let stale = stale_tag.checked_attributes();
    begin_attribute_census_0798();
    drop(stale);
    let tag = tag("e current=\"1\"");
    let mut current = tag.checked_attributes();
    assert!(current.next().unwrap().is_ok());
    assert!(current.next().is_none());
    drop(current);
    let report = finish_attribute_census_0798();
    assert_eq!(report.iterator_starts, 1);
    assert_eq!(report.iterator_drops, 1);
    assert_eq!(report.rows.len(), 1);
    assert_eq!(report.rows[0].element_name, b"e");
    assert_eq!(report.rows[0].source, b"e current=\"1\"");
}

#[test]
fn live_instances_are_reported_when_finish_precedes_drop() {
    begin_attribute_census_0798();
    let tag = tag("e a=\"1\"");
    let mut attributes = tag.checked_attributes();
    assert!(attributes.next().unwrap().is_ok());
    let report = finish_attribute_census_0798();
    assert_eq!(report.live_instances_at_finish, 1);
    assert_eq!(report.rows.len(), 1);
    assert!(report.rows[0].live_at_finish);
    assert!(!report.rows[0].dropped);
    assert!(report.rows[0].lexical_scan_completed);
    drop(attributes);
}

#[test]
fn foreign_thread_cannot_mutate_the_calling_thread_session() {
    begin_attribute_census_0798();
    let tag = Box::leak(Box::new(tag("e a=\"1\"")));
    let attributes = tag.checked_attributes();
    std::thread::spawn(move || drop(attributes))
        .join()
        .unwrap();
    let report = finish_attribute_census_0798();
    assert_eq!(report.live_instances_at_finish, 1);
    assert_eq!(report.iterator_drops, 0);
    assert!(report.rows[0].live_at_finish);
}

#[test]
fn counter_policy_saturates_without_wrapping() {
    let mut value = u64::MAX;
    let mut saturated = false;
    super::saturating_increment(&mut value, &mut saturated);
    assert_eq!(value, u64::MAX);
    assert!(saturated);
    super::saturating_add(&mut value, 1, &mut saturated);
    assert_eq!(value, u64::MAX);
    assert!(saturated);
}

#[test]
fn generation_overflow_is_reported_as_saturation() {
    super::ATTRIBUTE_CENSUS_GENERATION_0798.with(|serial| serial.set(u64::MAX));
    begin_attribute_census_0798();
    let report = finish_attribute_census_0798();
    assert!(report.counter_saturated);
    super::ATTRIBUTE_CENSUS_GENERATION_0798.with(|serial| serial.set(0));
}
