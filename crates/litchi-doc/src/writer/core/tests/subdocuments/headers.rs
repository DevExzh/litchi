use super::super::support::*;

#[test]
fn test_headers_and_footers() {
    let mut writer = Writer::new();
    writer.set_odd_header("Odd Header");
    writer.set_even_header("Even Header");
    writer.set_first_header("First Header");
    writer.set_odd_footer("Odd Footer");
    writer.set_even_footer("Even Footer");
    writer.set_first_footer("First Footer");
    assert_eq!(
        writer.header_odd.as_ref().unwrap()[0].runs[0].0,
        "Odd Header"
    );
    assert_eq!(
        writer.header_even.as_ref().unwrap()[0].runs[0].0,
        "Even Header"
    );
    assert_eq!(
        writer.header_first.as_ref().unwrap()[0].runs[0].0,
        "First Header"
    );
    assert_eq!(
        writer.footer_odd.as_ref().unwrap()[0].runs[0].0,
        "Odd Footer"
    );
    assert_eq!(
        writer.footer_even.as_ref().unwrap()[0].runs[0].0,
        "Even Footer"
    );
    assert_eq!(
        writer.footer_first.as_ref().unwrap()[0].runs[0].0,
        "First Footer"
    );
}

#[cfg(test)]
mod header_kind_tests {
    pub(super) use super::super::super::support::*;

    #[test]
    fn header_kinds_map_to_plcfhdd_slots() {
        assert_eq!(HeaderKind::Odd.slot(), HEADER_SLOT_ODD);
        assert_eq!(HeaderKind::Even.slot(), HEADER_SLOT_EVEN);
        assert_eq!(HeaderKind::FirstPage.slot(), HEADER_SLOT_FIRST);
        // The writer's slot assignment matches the MS-DOC PlcfHdd layout:
        // even header 6, odd header 7, first-page header 10.
        assert_eq!(
            (HEADER_SLOT_EVEN, HEADER_SLOT_ODD, HEADER_SLOT_FIRST),
            (6, 7, 10)
        );
    }

    #[test]
    fn header_shape_ids_use_the_header_cluster() {
        let mut writer = Writer::new();
        writer
            .insert_header_picture(
                HeaderKind::Odd,
                crate::writer::images::Picture::from_parts(
                    vec![0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A],
                    480,
                    240,
                )
                .unwrap(),
                crate::writer::images::FloatingPosition::new(0, 0),
            )
            .unwrap();
        writer
            .insert_header_text_box(
                HeaderKind::Even,
                crate::writer::shapes::Shape::new(
                    crate::writer::shapes::Kind::Rectangle,
                    1440,
                    720,
                )
                .unwrap(),
                crate::writer::images::FloatingPosition::new(0, 0),
                "box",
            )
            .unwrap();
        // One shared cluster for both kinds, in insertion order.
        assert_eq!(writer.header_pictures[0].shape_id, 2049);
        assert_eq!(writer.header_shapes[0].shape_id, 2050);
        // Anchors landed in the right header paragraph lists.
        assert_eq!(writer.header_odd.as_ref().unwrap().len(), 1);
        assert_eq!(writer.header_even.as_ref().unwrap().len(), 1);
        assert!(writer.header_first.is_none());
    }
}
