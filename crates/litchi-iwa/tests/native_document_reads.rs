//! Native chart and document reads after retiring unused reference traversal.

use litchi_iwa::{Document, application::Application, charts::Kind};

fn assert_native_reads(
    source: &[u8],
    application: Application,
    title: Option<&str>,
    rows: &[&str],
    expected_counts: (usize, usize, usize),
) {
    let document = Document::from_bytes(source).expect("saved native chart control");
    assert_eq!(document.application(), application);
    let charts = document.charts().expect("native chart metadata");
    assert_eq!(charts.len(), 1);
    let chart = &charts[0];
    assert_eq!(chart.title(), title);
    let row_names = chart
        .row_names()
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(row_names.as_slice(), rows);
    let column_names = chart
        .column_names()
        .iter()
        .map(String::as_str)
        .collect::<Vec<_>>();
    assert_eq!(column_names.as_slice(), ["April", "May", "June", "July"]);
    assert_eq!(chart.series_count(), 2);
    assert_eq!(chart.kind(), Kind::Column2d);
    assert!(!chart.contains_default_data());

    let stats = document.stats().expect("physical document statistics");
    assert_eq!(stats.application, application);
    assert_eq!(stats.total_objects, expected_counts.0);
    assert_eq!(stats.archives_count, expected_counts.1);
    assert_eq!(
        stats.message_type_counts.values().sum::<usize>(),
        expected_counts.2
    );
    assert_eq!(stats.message_type_counts.get(&5021), Some(&1));
    assert_eq!(stats.message_type_counts.get(&5023), Some(&1));

    let snapshot = document.snapshot();
    assert_eq!(snapshot.charts().unwrap(), charts);
    assert_eq!(
        snapshot.stats().unwrap().message_type_counts,
        stats.message_type_counts
    );
}

#[test]
fn native_pages_chart_and_statistics_survive_reference_graph_retirement() {
    assert_native_reads(
        include_bytes!("../../../test-data/iwork/pages/chart-arrangement-native.pages"),
        Application::Pages,
        Some("Pages Arrange native chart"),
        &["Region 1", "Region 2"],
        (581, 8, 587),
    );
}

#[test]
fn native_numbers_chart_and_statistics_survive_reference_graph_retirement() {
    assert_native_reads(
        include_bytes!("../../../test-data/iwork/numbers/chart-arrangement-native.numbers"),
        Application::Numbers,
        Some("Numbers Arrange native chart"),
        &["North", "South"],
        (635, 38, 644),
    );
}

#[test]
fn native_keynote_chart_and_statistics_survive_reference_graph_retirement() {
    assert_native_reads(
        include_bytes!("../../../test-data/iwork/keynote/chart-caption-native.key"),
        Application::Keynote,
        None,
        &["Region 1", "Region 2"],
        (966, 26, 972),
    );
}
