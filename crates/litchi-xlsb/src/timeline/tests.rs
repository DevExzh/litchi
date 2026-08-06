use super::{
    Cache, Filter, FilterType, Level, Range, State, View, Views, parse_cache, parse_views,
    write_cache, write_views,
};
use crate::package::error::Error;

fn cache() -> Cache {
    let bounds = Range::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z").unwrap();
    let mut state = State::new(bounds, 0, FilterType::DateBetween);
    state.selection = Some(Range::new("2024-02-01T00:00:00Z", "2024-03-01T00:00:00Z").unwrap());
    let mut cache = Cache::new("TimelineCache1", "Date", state);
    cache.filter = Some(Filter::new(2, 7));
    cache
}

#[test]
fn timeline_cache_xml_round_trip() {
    let original = cache();
    let bytes = write_cache(&original).unwrap();
    assert_eq!(parse_cache(&bytes).unwrap(), original);
}

#[test]
fn timeline_views_xml_round_trip() {
    let mut view = View::new("Timeline1", "TimelineCache1", Level::Month);
    view.selection_level = Level::Day;
    view.caption = Some("Date".to_string());
    view.show_header = false;
    view.scroll_position = Some("2024-06-01T00:00:00Z".to_string());
    view.style = Some("TimeSlicerStyleLight1".to_string());
    let views = Views { items: vec![view] };
    assert_eq!(parse_views(&write_views(&views).unwrap()).unwrap(), views);
}

#[test]
fn timeline_validation_rejects_inverted_ranges_and_unknown_filters() {
    assert!(Range::new("2025-01-01T00:00:00Z", "2024-01-01T00:00:00Z").is_err());
    assert!(matches!(
        parse_cache(
            br#"<x15:timelineCacheDefinition xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main" name="c" sourceName="d"><x15:state minimalRefreshVersion="0" lastRefreshVersion="0" pivotCacheId="0" filterType="nope"><x15:bounds startDate="2024-01-01T00:00:00Z" endDate="2024-01-02T00:00:00Z"/></x15:state></x15:timelineCacheDefinition>"#
        ),
        Err(Error::Unrecognized { .. })
    ));
}
