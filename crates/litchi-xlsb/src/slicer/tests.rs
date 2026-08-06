use super::{
    Cache, CrossFilter, Item, Native, SortOrder, View, Views, parse_cache, parse_views,
    write_cache, write_views,
};
use crate::package::error::Error;
use crate::raw::{Records, kind};

fn native_cache() -> Cache {
    let mut source = Native::new(0);
    source.sort_order = SortOrder::Descending;
    source.cross_filter = CrossFilter::ShowItemsWithoutData;
    source.items = vec![
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
    Cache::native("Slicer_State", "State", source)
}

#[test]
fn native_cache_round_trips_bounded_biff12() {
    let original = native_cache();
    let bytes = write_cache(&original).unwrap();
    let records: Vec<_> = Records::new(&bytes)
        .map(|record| record.unwrap().kind())
        .collect();
    assert_eq!(records[0], kind::BEGIN_SLICER_CACHE);
    assert_eq!(records.last().copied(), Some(kind::END_SLICER_CACHE));
    assert_eq!(parse_cache(&bytes).unwrap(), original);
}

#[test]
fn slicer_view_round_trips_optional_fields_and_flags() {
    let mut view = View::new("StateView", "Slicer_State");
    view.start_item = 2;
    view.column_count = 3;
    view.row_height = 228_600;
    view.caption_visible = false;
    view.caption = Some("State".to_string());
    view.style = Some("SlicerStyleLight1".to_string());
    view.locked_position = true;
    let views = Views { items: vec![view] };
    assert_eq!(parse_views(&write_views(&views).unwrap()).unwrap(), views);
}

#[test]
fn validation_rejects_unsafe_selection_and_reserved_bits() {
    let mut cache = native_cache();
    if let super::Source::Native(source) = &mut cache.source {
        source
            .items
            .iter_mut()
            .for_each(|item| item.selected = false);
    }
    assert!(matches!(
        write_cache(&cache),
        Err(Error::Unrecognized { .. })
    ));

    let mut bytes = write_cache(&native_cache()).unwrap();
    let mut records: Vec<(crate::raw::Kind, Vec<u8>)> = Records::new(&bytes)
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect();
    let native = records
        .iter_mut()
        .find(|(kind, _)| *kind == kind::BEGIN_SLICER_CACHE_NATIVE)
        .unwrap();
    native.1[8] |= 0x80;
    bytes.clear();
    let mut writer = crate::raw::Writer::new(&mut bytes);
    for (kind, payload) in records {
        writer.write_record(kind, &payload).unwrap();
    }
    assert!(parse_cache(&bytes).is_err());
}

#[test]
fn validation_rejects_zero_view_columns() {
    let mut views = Views {
        items: vec![View::new("StateView", "Slicer_State")],
    };
    views.items[0].column_count = 0;
    assert!(write_views(&views).is_err());
}
