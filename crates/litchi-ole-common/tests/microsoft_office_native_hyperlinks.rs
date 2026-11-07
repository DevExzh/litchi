//! Native Office reserved-property wire compatibility; targets remain inert.
use litchi_ole_common::property_set::{
    Binding, PropertySetReader, Stream, USER_DEFINED_PROPERTIES_FMTID, Value,
    user_defined::{Edit, HYPERLINKS, Limits, Properties},
};
use std::io::Cursor;

#[test]
fn native_excel_and_powerpoint_hyperlinks_use_one_blob_length_and_keep_producer_values() {
    let ppt_targets = [
        ("https://example.com/litchi-ppt-fixture", ""),
        ("", "260,5,DESTINATION_003"),
        (
            "mailto:fixture@example.com?subject=PPT_MAIL_SUBJECT_003",
            "",
        ),
        ("", "-1,-1,NEXT"),
        ("", "-1,-1,FIRST"),
    ];
    for (name, hash, app, info, targets) in [
        (
            "excel/excel-interaction-print.xls",
            65_629,
            0,
            6,
            vec![("https://example.com/litchi-excel-fixture", "")],
        ),
        (
            "powerpoint/powerpoint-navigation-review.ppt",
            7,
            6,
            7,
            ppt_targets.into_iter().chain(ppt_targets).collect(),
        ),
    ] {
        let bytes = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/microsoft-office")
                .join(name),
        )
        .unwrap();
        let mut ole =
            litchi_cfb::OleFile::open(Cursor::new(&bytes), litchi_cfb::OleFileLimits::default())
                .unwrap();
        let stream = ole
            .property_set(Binding::DocumentSummaryInformation)
            .unwrap();
        let mut section = stream
            .section(USER_DEFINED_PROPERTIES_FMTID)
            .unwrap()
            .clone();
        let Value::Blob(original) = section.find_named(HYPERLINKS).unwrap().1 else {
            panic!("BLOB")
        };
        let original = original.clone();
        assert_eq!(
            u32::from_le_bytes(original[..4].try_into().unwrap()),
            u32::try_from(targets.len() * 6).unwrap()
        );
        let links = Properties::new(&section, Limits::default())
            .unwrap()
            .hyperlinks()
            .unwrap()
            .unwrap();
        assert_eq!(links.len(), targets.len());
        for (link, (target, location)) in links.links().iter().zip(targets) {
            assert_eq!(link.stored_hash(), hash);
            assert_eq!(link.app(), app);
            assert_eq!(link.office_art(), 0);
            assert_eq!(link.info(), info);
            assert_eq!(link.target(), target);
            assert_eq!(link.location(), location);
        }
        let small_blob = Limits::builder()
            .max_blob_bytes(original.len() - 1)
            .build()
            .unwrap();
        assert!(matches!(
            Properties::new(&section, small_blob).unwrap().hyperlinks(),
            Err(litchi_ole_common::Error::Limit {
                kind: litchi_ole_common::LimitKind::UserDefinedBlobBytes,
                ..
            })
        ));
        if links.len() > 1 {
            let too_few = Limits::builder()
                .max_links(links.len() - 1)
                .build()
                .unwrap();
            assert!(matches!(
                Properties::new(&section, too_few).unwrap().hyperlinks(),
                Err(litchi_ole_common::Error::Limit {
                    kind: litchi_ole_common::LimitKind::UserDefinedLinks,
                    ..
                })
            ));
        }
        Edit::new(&mut section, Limits::default())
            .unwrap()
            .set_hyperlinks(links.clone())
            .unwrap();
        assert_eq!(
            section.find_named(HYPERLINKS).unwrap().1,
            &Value::Blob(original.clone())
        );
        let wire = Stream::new(section).to_bytes().unwrap();
        let mut marker = vec![0x41, 0, 0, 0];
        marker.extend_from_slice(&u32::try_from(original.len()).unwrap().to_le_bytes());
        marker.extend_from_slice(&original[..4]);
        assert!(wire.windows(marker.len()).any(|w| w == marker));
        let parsed = Stream::parse(&wire).unwrap();
        assert_eq!(
            Properties::new(
                parsed.section(USER_DEFINED_PROPERTIES_FMTID).unwrap(),
                Limits::default()
            )
            .unwrap()
            .hyperlinks()
            .unwrap(),
            Some(links)
        );
    }
}
