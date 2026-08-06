//! Transactional XLSB slicer/timeline package-owner regressions.

use litchi_xlsb::Package;
use litchi_xlsb::slicer::{
    Cache as SlicerCache, CrossFilter, Item, Native, View as SlicerView, Views as SlicerViews,
};
use litchi_xlsb::timeline::{
    Cache as TimelineCache, FilterType, Level, Range, State, View as TimelineView,
    Views as TimelineViews,
};
use std::io::Cursor;

fn slicer_cache() -> SlicerCache {
    let mut native = Native::new(0);
    native.cross_filter = CrossFilter::ShowItemsWithoutData;
    native.items = vec![
        Item {
            cache_index: 0,
            selected: true,
            no_data: false,
        },
        Item {
            cache_index: 1,
            selected: false,
            no_data: true,
        },
    ];
    SlicerCache::native("Slicer_State", "State", native)
}

fn timeline_cache() -> TimelineCache {
    let bounds = Range::new("2024-01-01T00:00:00Z", "2024-12-31T23:59:59Z").unwrap();
    TimelineCache::new(
        "TimelineCache1",
        "Date",
        State::new(bounds, 0, FilterType::DateBetween),
    )
}

#[test]
fn workbook_facade_round_trips_and_removes_both_inert_owners() {
    let mut workbook = Package::create().unwrap().into_workbook().unwrap();
    let cache = slicer_cache();
    workbook.set_slicer_caches(vec![cache.clone()]).unwrap();

    let mut slicer_view = SlicerView::new("StateView", "Slicer_State");
    slicer_view.caption = Some("State".to_string());
    workbook
        .set_slicers(
            0,
            SlicerViews {
                items: vec![slicer_view],
            },
        )
        .unwrap();

    let timeline = timeline_cache();
    workbook
        .set_timeline_caches(vec![timeline.clone()])
        .unwrap();
    workbook
        .set_timelines(
            0,
            TimelineViews {
                items: vec![TimelineView::new(
                    "Timeline1",
                    "TimelineCache1",
                    Level::Month,
                )],
            },
        )
        .unwrap();

    assert_eq!(workbook.slicer_caches().unwrap(), vec![cache.clone()]);
    assert_eq!(workbook.timeline_caches().unwrap(), vec![timeline.clone()]);
    assert!(workbook.slicers(0).unwrap().is_some());
    assert!(workbook.timelines(0).unwrap().is_some());

    let mut bytes = Cursor::new(Vec::new());
    workbook.save(&mut bytes).unwrap();
    let reopened = Package::from_bytes(bytes.into_inner())
        .unwrap()
        .into_workbook()
        .unwrap();
    assert_eq!(reopened.slicer_caches().unwrap(), vec![cache]);
    assert_eq!(reopened.timeline_caches().unwrap(), vec![timeline]);
    assert!(reopened.slicers(0).unwrap().is_some());
    assert!(reopened.timelines(0).unwrap().is_some());

    assert!(workbook.remove_slicers(0).unwrap());
    assert!(workbook.remove_slicer_caches().unwrap());
    assert!(workbook.remove_timelines(0).unwrap());
    assert!(workbook.remove_timeline_caches().unwrap());
    assert!(workbook.slicer_caches().unwrap().is_empty());
    assert!(workbook.timeline_caches().unwrap().is_empty());
    assert!(workbook.slicers(0).unwrap().is_none());
    assert!(workbook.timelines(0).unwrap().is_none());
}
