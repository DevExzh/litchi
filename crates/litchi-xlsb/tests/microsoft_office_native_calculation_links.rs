//! Native calculation-cache, hyperlink and protection metadata; all targets inert.
use litchi_xlsb::{
    Package,
    raw::{Limits, Records, kind},
};
use sha2::{Digest, Sha256};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/excel")
            .join(name),
    )
    .unwrap()
}
fn members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let a = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    a.member_names()
        .unwrap()
        .into_iter()
        .map(|n| {
            let b = a.read_member(&n).unwrap();
            (n, b)
        })
        .collect()
}
fn manifests(xml: &[u8], omit_chain: bool) -> Vec<BTreeMap<String, String>> {
    use quick_xml::events::Event;
    let mut reader = quick_xml::Reader::from_reader(xml);
    let mut out = Vec::new();
    loop {
        match reader.read_event().unwrap() {
            Event::Start(e) | Event::Empty(e)
                if matches!(
                    e.local_name().as_ref(),
                    b"Relationship" | b"Default" | b"Override"
                ) =>
            {
                let mut a = e
                    .attributes()
                    .map(|a| {
                        let a = a.unwrap();
                        (
                            String::from_utf8(a.key.as_ref().to_vec()).unwrap(),
                            a.decoded_and_normalized_value(
                                quick_xml::XmlVersion::Explicit1_0,
                                reader.decoder(),
                            )
                            .unwrap()
                            .into_owned(),
                        )
                    })
                    .collect::<BTreeMap<_, _>>();
                a.insert(
                    "element".into(),
                    String::from_utf8(e.local_name().as_ref().to_vec()).unwrap(),
                );
                if !omit_chain
                    || !a
                        .values()
                        .any(|s| s.ends_with("/calcChain") || s.ends_with("calcChain.bin"))
                {
                    out.push(a);
                }
            },
            Event::Eof => break,
            _ => {},
        }
    }
    out.sort();
    out
}
#[test]
fn native_calculation_properties_have_exact_flags_and_binary_payloads() {
    use litchi_xlsb::calc::{self, Mode, Opts};
    for name in [
        "excel-values-formulas-chart.xlsb",
        "excel-interaction-print.xlsb",
        "excel-metadata.xlsb",
        "excel-date1904-breaks.xlsb",
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let b = p.workbook().unwrap();
        let c = b.calc();
        assert_eq!(c.id(), 181029);
        assert_eq!(c.mode(), Mode::Auto);
        assert_eq!(c.iters(), 100);
        assert_eq!(c.delta().get(), 0.001);
        assert_eq!(c.threads().get(), 8);
        assert_eq!(
            c.opts(),
            Opts::A1 | Opts::FULL_PRECISION | Opts::RECALC_ON_SAVE | Opts::MTR
        );
        // Stored thread count is retained, but USER_THREADS is not set.
        let parts = members(&bytes);
        let raw = Records::new(&parts["xl/workbook.bin"], Limits::DEFAULT)
            .map(|r| r.unwrap())
            .filter(|r| r.kind() == kind::CALC_PROP)
            .map(|r| r.payload().to_vec())
            .collect::<Vec<_>>();
        assert_eq!(raw.len(), 1);
        assert_eq!(raw[0].len(), 26);
        assert_eq!(calc::read(&raw[0], Limits::DEFAULT).unwrap(), *c);
        let mut encoded = Vec::new();
        calc::write(
            c,
            &mut litchi_xlsb::raw::Encoder::new(&mut encoded, Limits::DEFAULT),
        )
        .unwrap();
        assert_eq!(encoded, raw[0]);
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
#[test]
fn native_calculation_chain_removal_preserves_workbook_and_reopens_exact_inverse() {
    for (stem, size, records, sha) in [
        (
            "excel-conditional-drawings",
            66,
            6,
            "057a67e18764416762e170880c0959ff38c503faa7691386dc8a7698f89fdabc",
        ),
        (
            "excel-date1904-breaks",
            51,
            5,
            "908ad0cee1f2e69612dd92d8de665b8162cd04f07033f4167a884bce4760d0c5",
        ),
        (
            "excel-formula-owners",
            1701,
            115,
            "44b51313997c95e777fdcab6a35589151b05665052758b355952326318119e61",
        ),
        (
            "excel-interaction-print",
            36,
            4,
            "52c8af5859c0cfc6131ec13025ca106a334290532e5d3049443038c0839b42db",
        ),
        (
            "excel-links-objects",
            51,
            5,
            "c2dbc0640df5f8464fa668d90fea7ecbbb1ee73321a1ff2002d5456b4863b544",
        ),
        (
            "excel-metadata",
            141,
            11,
            "d49a3d26eb038702a7faf8699b675d0259cbfd7ad339a2ed0d3b7d5aa48c153f",
        ),
        (
            "excel-pivot-controls",
            21,
            3,
            "549306349c364da76c94bac1714b3af94cc01ebb3b6d1ca806b5410e3b6bd39c",
        ),
        (
            "excel-revisions",
            21,
            3,
            "eba2b72d812fad5f675e0f207179bcc945228bbab14caeb3a78bca5b358e41f7",
        ),
        (
            "excel-values-formulas-chart",
            246,
            18,
            "d4678625f305c3b0719fbff836e93e86bea06a24214db7b5c44b460167b7a0ab",
        ),
        (
            "excel-visual-analytics",
            21,
            3,
            "549306349c364da76c94bac1714b3af94cc01ebb3b6d1ca806b5410e3b6bd39c",
        ),
    ] {
        let bytes = native(&format!("{stem}.xlsb"));
        let p = Package::from_bytes(&bytes).unwrap();
        let snapshot = p.calculation_chain().unwrap();
        let part = snapshot.part().unwrap();
        assert_eq!(part.part_name(), "/xl/calcChain.bin");
        assert_eq!(
            part.content_type(),
            litchi_xlsb::calculation_chain::CONTENT_TYPE
        );
        assert_eq!(part.len(), size);
        assert_eq!(
            Sha256::digest(part.bytes())
                .iter()
                .map(|b| format!("{b:02x}"))
                .collect::<String>(),
            sha
        );
        assert_eq!(part.record_count().unwrap(), records); // framing diagnostic, not a typed chain grammar
        assert!(
            p.edit_calculation_chain()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap()
                == bytes
        );
        let mut edit = p.edit_calculation_chain().unwrap();
        assert!(edit.remove().unwrap());
        assert!(!edit.remove().unwrap());
        let commit = edit.commit().unwrap();
        let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
        commit.package().save(f.path()).unwrap();
        let after = std::fs::read(f.path()).unwrap();
        let read = Package::from_bytes(&after).unwrap();
        assert!(!read.calculation_chain().unwrap().is_present());
        assert_eq!(
            read.workbook().unwrap().calc(),
            p.workbook().unwrap().calc()
        );
        let a = members(&bytes);
        let b = members(&after);
        assert_eq!(
            a.keys()
                .filter(|n| n.as_str() != "xl/calcChain.bin")
                .collect::<Vec<_>>(),
            b.keys().collect::<Vec<_>>()
        );
        for (n, data) in &a {
            match n.as_str() {
                "xl/calcChain.bin" => {},
                "[Content_Types].xml" | "xl/_rels/workbook.bin.rels" => {
                    assert_eq!(manifests(data, true), manifests(&b[n], true));
                    assert!(!std::str::from_utf8(&b[n]).unwrap().contains("calcChain"));
                },
                _ => assert!(b[n] == *data, "{stem} changed {n}"),
            }
        }
        let restored = read
            .apply_calculation_chain(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap();
        let restored = members(&restored);
        assert_eq!(
            restored.keys().collect::<Vec<_>>(),
            a.keys().collect::<Vec<_>>()
        );
        for (n, data) in a {
            assert!(restored[&n] == data, "{stem} inverse changed {n}");
        }
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
#[test]
fn native_chain_record_limit_does_not_block_opaque_removal() {
    use litchi_xlsb::calculation_chain::{LimitKind, SemanticLimits};
    let bytes = native("excel-values-formulas-chart.xlsb");
    let p = Package::from_bytes(&bytes).unwrap().with_semantic_limits(
        litchi_xlsb::SemanticLimits::default().calculation_chain(SemanticLimits {
            max_records: 17,
            ..SemanticLimits::DEFAULT
        }),
    );
    let s = p.calculation_chain().unwrap();
    let part = s.part().unwrap();
    assert_eq!(part.len(), 246);
    assert!(matches!(
        part.record_count(),
        Err(litchi_xlsb::package::PackageError::Limit {
            kind: litchi_xlsb::package::LimitKind::CalculationChain(LimitKind::Records),
            ..
        })
    ));
    let mut edit = p.edit_calculation_chain().unwrap();
    assert!(edit.remove().unwrap());
    let c = edit.commit().unwrap();
    assert!(!c.package().calculation_chain().unwrap().is_present());
    assert!(
        c.package()
            .apply_calculation_chain(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
#[test]
fn native_hyperlink_and_sheet_protection_agree_across_read_routes() {
    let bytes = native("excel-interaction-print.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let links = p.edit_hyperlinks().unwrap().hyperlinks(0).unwrap();
    assert_eq!(links.len(), 1);
    let l = &links[0];
    assert_eq!(
        (l.row_first, l.row_last, l.col_first, l.col_last),
        (9, 9, 0, 0)
    );
    assert_eq!(l.r_id, "rId2");
    assert_eq!(
        l.target.as_deref(),
        Some("https://example.com/litchi-excel-fixture")
    );
    assert_eq!(
        l.tooltip.as_deref(),
        Some("Native hyperlink screen tip 002")
    );
    assert!(l.location.is_none() && l.display.is_none());
    let lazy = litchi_xlsb::SourceBackedWorkbook::from_reader(Cursor::new(&bytes)).unwrap();
    let features = lazy
        .worksheet_by_name("Sheet1")
        .unwrap()
        .unwrap()
        .features(litchi_xlsb::sheet::FeatureSet::HYPERLINKS)
        .unwrap();
    assert_eq!(features.hyperlinks, links);
    let b = p.workbook().unwrap();
    let sheet = b.worksheet(0).unwrap();
    let protection = sheet.sheet_protection().unwrap();
    assert!(protection.locked && protection.allow_sort && protection.allow_auto_filter);
    assert!(protection.allow_select_locked_cells && protection.allow_select_unlocked_cells);
    assert_eq!(protection.password_hash, None);
    assert!(sheet.strong_sheet_protection().is_none());
    assert!(
        !protection.allow_edit_objects
            && !protection.allow_edit_scenarios
            && !protection.allow_format_cells
            && !protection.allow_format_columns
            && !protection.allow_format_rows
            && !protection.allow_insert_columns
            && !protection.allow_insert_rows
            && !protection.allow_insert_hyperlinks
            && !protection.allow_delete_columns
            && !protection.allow_delete_rows
            && !protection.allow_pivot_tables
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
fn records_except_links(data: &[u8]) -> Vec<Vec<u8>> {
    let records = Records::new(data, Limits::DEFAULT)
        .map(|r| r.unwrap())
        .collect::<Vec<_>>();
    records
        .iter()
        .enumerate()
        .filter(|(_, r)| r.kind() != kind::H_LINK)
        .map(|(index, r)| {
            let end = records
                .get(index + 1)
                .map_or(data.len(), |next| next.offset());
            data[r.offset()..end].to_vec()
        })
        .collect()
}

#[test]
fn native_hyperlink_retarget_add_remove_preserves_cells_protection_and_cache() {
    use litchi_xlsb::hyperlinks::Hyperlink;
    let bytes = native("excel-interaction-print.xlsb");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_hyperlinks().unwrap();
    let mut link = edit.hyperlinks(0).unwrap()[0].clone();
    assert!(!edit.set(0, 0, link.clone()).unwrap());
    link.target = Some("https://example.com/litchi-reviewed".into());
    link.tooltip = Some("Reviewed native link".into());
    assert!(edit.set(0, 0, link).unwrap());
    assert_eq!(
        edit.add(0, Hyperlink::new_internal(9, 9, 1, 1, "Sheet1!A1".into()))
            .unwrap(),
        1
    );
    let c = edit.commit().unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".xlsb").unwrap();
    c.package().save(f.path()).unwrap();
    let after = std::fs::read(f.path()).unwrap();
    let read = Package::from_bytes(&after).unwrap();
    let links = read.edit_hyperlinks().unwrap().hyperlinks(0).unwrap();
    assert_eq!(links.len(), 2);
    assert_eq!(
        links[0].target.as_deref(),
        Some("https://example.com/litchi-reviewed")
    );
    assert_eq!(links[0].tooltip.as_deref(), Some("Reviewed native link"));
    assert_eq!(links[1].location.as_deref(), Some("Sheet1!A1"));
    let a = members(&bytes);
    let b = members(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    assert_eq!(
        a.iter()
            .filter(|(n, v)| b[*n] != **v)
            .map(|(n, _)| n.as_str())
            .collect::<Vec<_>>(),
        [
            "xl/worksheets/_rels/sheet1.bin.rels",
            "xl/worksheets/sheet1.bin"
        ]
    );
    assert_eq!(
        records_except_links(&a["xl/worksheets/sheet1.bin"]),
        records_except_links(&b["xl/worksheets/sheet1.bin"])
    );
    let old_rels = manifests(&a["xl/worksheets/_rels/sheet1.bin.rels"], false);
    let new_rels = manifests(&b["xl/worksheets/_rels/sheet1.bin.rels"], false);
    assert_eq!(old_rels.len(), new_rels.len());
    let outside_links = |rels: &[BTreeMap<String, String>]| {
        rels.iter()
            .filter(|r| !r["Type"].ends_with("/hyperlink"))
            .cloned()
            .collect::<Vec<_>>()
    };
    assert_eq!(outside_links(&old_rels), outside_links(&new_rels));
    let external = new_rels
        .iter()
        .filter(|r| r["Type"].ends_with("/hyperlink"))
        .collect::<Vec<_>>();
    assert_eq!(external.len(), 1);
    assert_eq!(external[0]["Id"], links[0].r_id);
    assert_eq!(external[0]["Target"], "https://example.com/litchi-reviewed");
    assert_eq!(external[0]["TargetMode"], "External");
    let inverse = read
        .apply_hyperlinks(&c.patch().inverse())
        .unwrap()
        .package()
        .to_bytes()
        .unwrap();
    let restored = members(&inverse);
    assert_eq!(
        restored.keys().collect::<Vec<_>>(),
        a.keys().collect::<Vec<_>>()
    );
    for (n, data) in &a {
        assert!(restored[n] == *data, "inverse changed {n}");
    }
    let mut remove = read.edit_hyperlinks().unwrap();
    remove.remove(0, 1).unwrap();
    remove.remove(0, 0).unwrap();
    let removed = remove.commit().unwrap();
    let removed_parts = members(&removed.package().to_bytes().unwrap());
    assert_eq!(
        manifests(&removed_parts["xl/worksheets/_rels/sheet1.bin.rels"], false),
        outside_links(&old_rels)
    );
    assert_eq!(
        records_except_links(&removed_parts["xl/worksheets/sheet1.bin"]),
        records_except_links(&a["xl/worksheets/sheet1.bin"])
    );
    assert!(
        removed
            .package()
            .edit_hyperlinks()
            .unwrap()
            .hyperlinks(0)
            .unwrap()
            .is_empty()
    );
    assert!(
        removed
            .package()
            .apply_hyperlinks(&removed.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == after
    );
}
#[test]
fn native_invalid_hyperlink_replacement_preserves_the_staged_owner() {
    use litchi_xlsb::hyperlinks::Hyperlink;
    let bytes = native("excel-interaction-print.xlsb");
    let p = Package::from_bytes_with_options(
        &bytes,
        litchi_xlsb::OpenOptions::new().raw_limits(Limits::new(64 * 1024 * 1024, 256)),
    )
    .unwrap();
    let mut edit = p.edit_hyperlinks().unwrap();
    let original = edit.hyperlinks(0).unwrap();
    let mut invalid = original[0].clone();
    invalid.row_first = 10;
    invalid.row_last = 9;
    assert!(matches!(
        edit.set(0, 0, invalid),
        Err(litchi_xlsb::package::PackageError::InvalidFormat(
            "the hyperlink range is outside the worksheet grid"
        ))
    ));
    assert_eq!(edit.hyperlinks(0).unwrap(), original);
    assert!(edit.clone().commit().unwrap().package().to_bytes().unwrap() == bytes);
    edit.add(0, Hyperlink::new_internal(9, 9, 1, 1, "Sheet1!A1".into()))
        .unwrap();
    let expected = edit.clone().commit().unwrap().package().to_bytes().unwrap();
    let staged = edit.hyperlinks(0).unwrap();
    for replace in [false, true] {
        let mut invalid = original[0].clone();
        invalid.target = Some("https://example.com/rejected".into());
        invalid.tooltip = Some("x".repeat(257));
        let result = if replace {
            edit.set(0, 0, invalid).map(|_| ())
        } else {
            edit.add(0, invalid).map(|_| ())
        };
        assert!(matches!(
            result,
            Err(litchi_xlsb::package::PackageError::Hyperlink(
                litchi_xlsb::hyperlinks::Error::Wire(litchi_xlsb::raw::Error::StringLimit {
                    units: 257,
                    limit: 256,
                    ..
                })
            ))
        ));
        assert_eq!(edit.hyperlinks(0).unwrap(), staged);
        assert!(edit.clone().commit().unwrap().package().to_bytes().unwrap() == expected);
    }
    assert!(matches!(
        edit.remove(0, usize::MAX),
        Err(litchi_xlsb::package::PackageError::InvalidFormat(
            "the hyperlink is missing"
        ))
    ));
    assert_eq!(edit.hyperlinks(0).unwrap(), staged);
    assert!(edit.commit().unwrap().package().to_bytes().unwrap() == expected);
    assert!(p.to_bytes().unwrap() == bytes);
}
