//! PivotTable codec and aggregate tests.

use super::*;

#[cfg(test)]
mod worksheet_view_record_tests {
    use super::*;

    fn view_payload() -> Vec<u8> {
        let mut data = vec![0u8; 44];
        data[2..4].copy_from_slice(&1u16.to_le_bytes());
        data[6..8].copy_from_slice(&1u16.to_le_bytes());
        data[10..12].copy_from_slice(&1u16.to_le_bytes());
        data[12..14].copy_from_slice(&1u16.to_le_bytes());
        data[36..38].copy_from_slice(&0x020Bu16.to_le_bytes());
        data[38..40].copy_from_slice(&1u16.to_le_bytes());
        data[40..42].copy_from_slice(&1u16.to_le_bytes());
        data[42..44].copy_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&[0, b'P', 0, b'V']);
        data
    }

    fn sxex_payload() -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        for _ in 0..3 {
            data.extend_from_slice(&u16::MAX.to_le_bytes());
        }
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&0x004F_0200u32.to_le_bytes());
        for _ in 0..3 {
            data.extend_from_slice(&u16::MAX.to_le_bytes());
        }
        data
    }

    #[test]
    fn parses_current_sxview_layout_and_cache_index() {
        let mut payload = view_payload();
        payload[14..16].copy_from_slice(&7u16.to_le_bytes());
        let view = parse_sxview(&payload).unwrap();
        assert_eq!(view.cache_index, 7);
        assert_eq!(view.name, "P");
        assert_eq!(view.data_field_name, "V");
    }

    #[test]
    fn rejects_out_of_order_and_duplicate_singletons() {
        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        assert!(collector.feed_record(SXLI_TYPE, &[]).is_err());

        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        collector.feed_record(SXEX_TYPE, &sxex_payload()).unwrap();
        assert!(collector.feed_record(SXEX_TYPE, &sxex_payload()).is_err());
    }

    #[test]
    fn preserves_sxaddl_payload_exactly_and_rejects_bad_lengths() {
        let mut collector = PivotTableCollector::new();
        collector.feed_record(SXVIEW_TYPE, &view_payload()).unwrap();
        collector.feed_record(SXEX_TYPE, &sxex_payload()).unwrap();
        let payload = [0x64, 0x08, 0, 0, 0, 2, 0xAA, 0xBB, 0xCC];
        collector.feed_record(SXADDL_TYPE, &payload).unwrap();
        let tables = collector.finish().unwrap();
        assert_eq!(
            tables[0].additional_extensions[0].payload,
            [0xAA, 0xBB, 0xCC]
        );
        assert!(parse_sxaddl(&payload[..5]).is_err());
        assert!(parse_sxivd(&[0]).is_err());
        assert!(parse_sxpi(&[0; 5]).is_err());
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_sxvs() {
        let data = 0x0001u16.to_le_bytes();
        assert_eq!(parse_sxvs(&data).unwrap(), PivotSourceType::Worksheet);
    }

    #[test]
    fn test_parse_sxpi_two_entries() {
        let mut data = Vec::new();
        // Entry 1
        data.extend_from_slice(&0u16.to_le_bytes()); // isxvd
        data.extend_from_slice(&1u16.to_le_bytes()); // isxvi
        data.extend_from_slice(&0u16.to_le_bytes()); // idObj
        // Entry 2
        data.extend_from_slice(&1u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&0u16.to_le_bytes());

        let entries = parse_sxpi(&data).unwrap();
        assert_eq!(entries.len(), 2);
        assert_eq!(entries[0].item_index, 1);
        assert_eq!(entries[1].field_index, 1);
    }

    #[test]
    fn test_parse_sxvd_no_name() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x0001u16.to_le_bytes()); // axis = Row
        data.extend_from_slice(&0u16.to_le_bytes()); // cSub
        data.extend_from_slice(&0u16.to_le_bytes()); // grbitSub
        data.extend_from_slice(&5u16.to_le_bytes()); // cItm
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // cchName = not present

        let field = parse_sxvd(&data).unwrap();
        assert_eq!(field.axis, PivotAxis::Row);
        assert_eq!(field.item_count, 5);
        assert!(field.name.is_none());
    }

    #[test]
    fn test_parse_sxvi_data_item() {
        let mut data = Vec::new();
        data.extend_from_slice(&0x00FEu16.to_le_bytes()); // itmType = Data
        data.extend_from_slice(&0u16.to_le_bytes()); // flags
        data.extend_from_slice(&3u16.to_le_bytes()); // iCache
        data.extend_from_slice(&0xFFFFu16.to_le_bytes()); // no name

        let item = parse_sxvi(&data).unwrap();
        assert_eq!(item.item_type, PivotItemType::Data);
        assert_eq!(item.cache_index, 3);
        assert!(item.name.is_none());
    }
}
