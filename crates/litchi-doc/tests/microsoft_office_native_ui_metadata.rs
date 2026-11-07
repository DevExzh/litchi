//! Native command tables and inert routing metadata; no command or mail execution.
use litchi_doc::parts::command_bars as bars;
use litchi_doc::route_slip as route;
use litchi_doc::{Package, SourceLimits};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
const MACROS: [&str; 2] = ["word-recorded-macro.doc", "word-recorded-macro.dot"];
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let bytes = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, bytes)
        })
        .collect()
}
fn table_name(p: &Package) -> Vec<String> {
    vec![format!(
        "{}Table",
        u8::from(p.document().fib().which_table_stream())
    )]
}
fn payload(p: &Package, index: usize) -> Vec<u8> {
    let (at, len) = p.document().fib().get_table_pointer(index).unwrap();
    let data = streams(&p.to_bytes().unwrap());
    data[&table_name(p)][at as usize..(at + len) as usize].to_vec()
}
fn command_bars(p: &Package) -> Option<bars::CommandBars<'static>> {
    p.edit_command_bars()
        .unwrap()
        .command_bars()
        .unwrap()
        .command_bars()
        .cloned()
}
fn preserved(p: &Package, q: &Package, index: usize) {
    let before = p.to_bytes().unwrap();
    let after = q.to_bytes().unwrap();
    assert_eq!(
        cfb_metadata::directory(&before),
        cfb_metadata::directory(&after)
    );
    assert_eq!(p.document().text().unwrap(), q.document().text().unwrap());
    let a = streams(&before);
    let b = streams(&after);
    let table = table_name(p);
    assert_eq!(table, table_name(q));
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in &a {
        if *name != table && name != &["WordDocument"] {
            assert!(b.get(name) == Some(bytes), "changed {}", name.join("/"));
        }
    }
    assert!(b[&table].starts_with(&a[&table]));
    let word = vec!["WordDocument".into()];
    let pair = 154 + index * 8;
    assert_eq!(&a[&word][..pair], &b[&word][..pair]);
    assert_eq!(&a[&word][pair + 8..], &b[&word][pair + 8..]);
}
fn reopen(p: &Package) -> Package {
    let file = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    p.save(file.path()).unwrap();
    let q = Package::open(file.path()).unwrap();
    assert_eq!(q.to_bytes().unwrap(), p.to_bytes().unwrap());
    q
}
fn bars_publish(p: &Package, c: &bars::Commit) -> Package {
    let q = reopen(c.package());
    preserved(p, &q, bars::FIB_INDEX_CMDS);
    assert_eq!(
        p.apply_command_bars(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        q.to_bytes().unwrap()
    );
    assert_eq!(
        q.apply_command_bars(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        p.to_bytes().unwrap()
    );
    q
}
fn route_publish(p: &Package, c: &route::Commit) -> Package {
    let q = reopen(c.package());
    preserved(p, &q, route::FIB_INDEX_ROUTE_SLIP);
    assert_eq!(
        p.apply_route_slip(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        q.to_bytes().unwrap()
    );
    assert_eq!(
        q.apply_route_slip(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        p.to_bytes().unwrap()
    );
    q
}
fn expected_native_commands() -> bars::CommandBars<'static> {
    use bars::*;
    CommandBars::new(vec![
        Entry::MacroCommands(
            MacroCommands::new(vec![MacroCommand::new(0, 0, 0x56, 0, 0xffff, 0, 0, 0, 0)]).unwrap(),
        ),
        Entry::CommandStrings(
            CommandStrings::new(vec![
                CommandString::new(
                    XString::new("Project.NewMacros.FixtureBoldSelection").unwrap(),
                    1,
                )
                .unwrap(),
            ])
            .unwrap(),
        ),
        Entry::MacroNames(
            MacroNames::new(vec![
                MacroName::new(
                    0,
                    XString::new("PROJECT.NEWMACROS.FIXTUREBOLDSELECTION").unwrap(),
                )
                .unwrap(),
            ])
            .unwrap(),
        ),
    ])
    .unwrap()
}
fn extended(mut value: bars::CommandBars<'static>) -> bars::CommandBars<'static> {
    use bars::*;
    use litchi_ole_common::toolbar::{
        ControlFlags, ControlHeader, ControlType, Data, Flags, GeneralFlags, GeneralInfo, Header,
        Restrictions, SpecificFlags, Type, WString,
    };
    let control = Control::new(
        ControlHeader::new(
            ControlType::Label,
            0x1234,
            ControlFlags::from_raw(0),
            SpecificFlags::from_raw(0),
            0,
            None,
        )
        .unwrap(),
        Some(CommandId::new(0x1231).unwrap()),
        Some(
            Data::new(
                GeneralInfo::new(GeneralFlags::default(), None, None, None, None).unwrap(),
                Vec::new(),
            )
            .unwrap(),
        ),
    )
    .unwrap();
    let toolbar = Toolbar::new(
        XString::new("Native toolbar 日本語 🍎").unwrap(),
        Header::new(
            0,
            Restrictions::new(Type::Basic).unwrap(),
            0,
            Flags::default(),
            WString::new("Native controls").unwrap(),
        )
        .unwrap(),
        [0xa5; 100],
        0,
        0,
        0xcafe,
    )
    .unwrap()
    .with_controls(vec![control])
    .unwrap();
    value.entries_mut().push(Entry::KeyMaps(
        KeyMaps::new(
            KeyMapKind::Regular,
            vec![KeyMap::new(0x241, 0xff, Action::Character, 0x3042, 0, 0)],
        )
        .unwrap(),
    ));
    value.entries_mut().push(Entry::Toolbar(
        ToolbarWrapper::new(vec![Customization::toolbar(toolbar).unwrap()]).unwrap(),
    ));
    value.validate().unwrap();
    value
}

#[test]
fn native_doc_macro_command_tables_have_exact_names_descriptors_and_wire_bytes() {
    for file in MACROS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let value = command_bars(&p).unwrap();
        assert_eq!(value, expected_native_commands());
        assert_eq!(payload(&p, bars::FIB_INDEX_CMDS).len(), 203);
        assert_eq!(
            bars::to_bytes(&value, &SourceLimits::default()).unwrap(),
            payload(&p, bars::FIB_INDEX_CMDS)
        );
        let mut e = p.edit_command_bars().unwrap();
        e.replace(value).unwrap();
        let c = e.commit().unwrap();
        assert!(c.patch().is_noop());
        assert_eq!(c.package().to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
    let p = Package::from_bytes(&native("word-text-structure.doc")).unwrap();
    assert!(command_bars(&p).unwrap().entries().is_empty());
    assert_eq!(payload(&p, bars::FIB_INDEX_CMDS), [0xff, 0x40]);
}

#[test]
fn native_doc_command_keymap_toolbar_lifecycle_preserves_vba_and_other_tables() {
    use bars::Entry;
    for file in MACROS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let value = command_bars(&p).unwrap();
        let mut expected = extended(value.clone());
        let mut e = p.edit_command_bars().unwrap();
        e.replace(expected.clone()).unwrap();
        let c = e.commit().unwrap();
        let added = bars_publish(&p, &c);
        assert_eq!(command_bars(&added).unwrap(), expected);
        assert_eq!(&expected.entries()[..3], value.entries());
        if let Entry::CommandStrings(strings) = &mut expected.entries_mut()[1] {
            strings.strings_mut()[0] = bars::CommandString::new(
                bars::XString::new("Project.NewMacros.FixtureBoldSelection").unwrap(),
                2,
            )
            .unwrap();
        } else {
            panic!("native command strings");
        }
        let mut e = added.edit_command_bars().unwrap();
        e.replace(expected.clone()).unwrap();
        let mut invalid = expected.clone();
        invalid
            .entries_mut()
            .retain(|e| !matches!(e, Entry::MacroNames(_)));
        assert!(
            matches!(e.replace(invalid),Err(bars::Error::Package(litchi_doc::Error::Corrupted(m))) if m=="PlfMcd requires a MacroNames table for its macro-name indexes")
        );
        assert_eq!(e.command_bars().unwrap().command_bars(), Some(&expected));
        let c = e.commit().unwrap();
        let updated = bars_publish(&added, &c);
        assert_eq!(command_bars(&updated).unwrap(), expected);
        assert!(matches!(
            updated.apply_command_bars(c.patch()),
            Err(bars::Error::PatchConflict)
        ));
        let mut e = updated.edit_command_bars().unwrap();
        e.clear().unwrap();
        let c = e.commit().unwrap();
        let removed = bars_publish(&updated, &c);
        assert!(command_bars(&removed).is_none());
        assert_eq!(
            removed
                .document()
                .fib()
                .get_table_pointer(bars::FIB_INDEX_CMDS)
                .unwrap(),
            (0, 0)
        );
        assert_eq!(
            streams(&removed.to_bytes().unwrap())[&table_name(&removed)],
            streams(&updated.to_bytes().unwrap())[&table_name(&updated)]
        );
        let mut e = removed.edit_command_bars().unwrap();
        e.set(value.clone()).unwrap();
        let c = e.commit().unwrap();
        let recreated = bars_publish(&removed, &c);
        assert_eq!(command_bars(&recreated), Some(value));
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

fn routing() -> route::Metadata {
    use route::{DeliveryOption, Metadata, NarrowString as Text, Protection, Recipient};
    Metadata::try_new(
        true,
        false,
        true,
        Protection::Off,
        0,
        DeliveryOption::Serial,
        Text::new(vec![0x80, b's', b'u', b'b']),
        Text::new(b"Stored message".to_vec()),
        Text::new(vec![0x81, b'o', b'k']),
        Text::new(b"Native routing metadata".to_vec()),
        vec![
            Recipient::try_new(
                vec![0, 0xff, 1],
                Text::new(b"first@example.invalid".to_vec()),
            )
            .unwrap(),
            Recipient::try_new(
                vec![2, 0, 3],
                Text::new(vec![0x90, b's', b'e', b'c', b'o', b'n', b'd']),
            )
            .unwrap(),
        ],
    )
    .unwrap()
}

#[test]
fn native_doc_routing_lifecycle_preserves_narrow_bytes_and_unowned_document_state() {
    use route::{NarrowString, Recipient, RecipientSelector as Select};
    for file in ["word-text-structure.doc", MACROS[0], MACROS[1]] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.document().route_slip().unwrap().is_none());
        let mut expected = routing();
        let mut e = p.edit_route_slip().unwrap();
        e.set(expected.clone()).unwrap();
        let c = e.commit().unwrap();
        let added = route_publish(&p, &c);
        assert_eq!(added.document().route_slip().unwrap(), Some(&expected));
        assert_eq!(
            payload(&added, route::FIB_INDEX_ROUTE_SLIP),
            expected.to_bytes(&SourceLimits::default()).unwrap()
        );
        let mut e = added.edit_route_slip().unwrap();
        e.replace(expected.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_route_slip().unwrap();
        e.advance_stage().unwrap();
        expected.stage = 1;
        assert!(
            matches!(e.advance_stage(),Err(route::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m=="route-slip transaction failed: Corrupted file: route slip is already at its final recipient")
        );
        assert_eq!(e.metadata().unwrap(), Some(&expected));
        let third = Recipient::try_new(
            vec![0x76, 0xff],
            NarrowString::new(b"third@example.invalid".to_vec()),
        )
        .unwrap();
        e.add_recipient(third.clone()).unwrap();
        expected.recipients.push(third);
        e.set_stage(Select::Name(b"third@example.invalid")).unwrap();
        expected.stage = 2;
        let c = e.commit().unwrap();
        let updated = route_publish(&added, &c);
        assert_eq!(updated.document().route_slip().unwrap(), Some(&expected));
        assert!(matches!(
            updated.apply_route_slip(c.patch()),
            Err(route::Error::PatchConflict)
        ));
        let mut e = updated.edit_route_slip().unwrap();
        e.remove_recipient(Select::Index(0)).unwrap();
        expected.recipients.remove(0);
        expected.stage = 1;
        let c = e.commit().unwrap();
        let shortened = route_publish(&updated, &c);
        assert_eq!(shortened.document().route_slip().unwrap(), Some(&expected));
        let mut e = shortened.edit_route_slip().unwrap();
        e.complete().unwrap();
        let c = e.commit().unwrap();
        let cleared = route_publish(&shortened, &c);
        assert!(cleared.document().route_slip().unwrap().is_none());
        assert_eq!(
            cleared
                .document()
                .fib()
                .get_table_pointer(route::FIB_INDEX_ROUTE_SLIP)
                .unwrap(),
            (0, 0)
        );
        assert_eq!(
            streams(&cleared.to_bytes().unwrap())[&table_name(&cleared)],
            streams(&shortened.to_bytes().unwrap())[&table_name(&shortened)]
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_doc_route_protection_applies_to_replacement_and_recipient_edits() {
    use route::{EditKind, Protection, RecipientSelector};
    let bytes = native("word-text-structure.doc");
    let p = Package::from_bytes(&bytes).unwrap();
    for (protection, allowed) in [
        (Protection::RevisionMark, [false, true, false, false]),
        (Protection::Annotation, [false, false, true, false]),
        (Protection::Form, [false, false, false, true]),
    ] {
        assert_eq!(
            [
                EditKind::Content,
                EditKind::Revision,
                EditKind::Annotation,
                EditKind::FormField
            ]
            .map(|k| protection.allows(k)),
            allowed
        );
        let mut value = routing();
        value.protection = protection;
        let mut e = p.edit_route_slip().unwrap();
        e.set(value.clone()).unwrap();
        let c = e.commit().unwrap();
        let protected = route_publish(&p, &c);
        let mut e = protected.edit_route_slip().unwrap();
        e.replace(value.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = protected.edit_route_slip().unwrap();
        let mut replacement = value.clone();
        replacement.protection = Protection::Off;
        let reason = format!(
            "route-slip transaction failed: route-slip metadata is protected by {protection:?}"
        );
        assert!(
            matches!(e.set(replacement.clone()),Err(route::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m==reason)
        );
        assert!(
            matches!(e.replace(replacement),Err(route::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m==reason)
        );
        assert!(
            matches!(e.set_stage(RecipientSelector::Index(0)),Err(route::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m==reason)
        );
        assert!(
            matches!(e.clear(),Err(route::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m==reason)
        );
        assert_eq!(e.metadata().unwrap(), Some(&value));
        assert_eq!(
            e.commit().unwrap().package().to_bytes().unwrap(),
            protected.to_bytes().unwrap()
        );
    }
    assert_eq!(p.to_bytes().unwrap(), bytes);
    assert_eq!(native("word-text-structure.doc"), bytes);
}

// Deliberately non-signing sentinel payloads. Only the MS-OSHARED container
// framing is valid; these bytes establish no PKCS#7 validity or trust.
fn signature_blob(word: bool) -> Vec<u8> {
    let mut bytes = vec![0; 56];
    let values = if word {
        vec![
            (2, 45),
            (6, 8),
            (10, 2),
            (14, 44),
            (18, 3),
            (22, 46),
            (30, 49),
            (42, 51),
        ]
    } else {
        vec![
            (0, 48),
            (4, 8),
            (8, 3),
            (12, 44),
            (16, 4),
            (20, 48),
            (28, 52),
            (32, 0xa5a5_5a5a),
            (40, 54),
        ]
    };
    for (offset, value) in values {
        bytes[offset..offset + 4].copy_from_slice(&u32::to_le_bytes(value));
    }
    if word {
        bytes[..2].copy_from_slice(&27u16.to_le_bytes());
        bytes[46..48].copy_from_slice(&[9, 8]);
        bytes[48..51].copy_from_slice(&[7, 6, 5]);
        bytes[55] = 0xee;
    } else {
        bytes[44..47].copy_from_slice(&[1, 2, 3]);
        bytes[47] = 0xcc;
        bytes[48..52].copy_from_slice(&[4, 5, 6, 7]);
    }
    bytes
}
fn properties_publish(p: &Package, c: &litchi_doc::package::property_set::Commit) -> Package {
    use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
    let q = reopen(c.package());
    let before = p.to_bytes().unwrap();
    let after = q.to_bytes().unwrap();
    assert_eq!(
        cfb_metadata::directory(&before),
        cfb_metadata::directory(&after)
    );
    assert_eq!(p.document().text().unwrap(), q.document().text().unwrap());
    let a = streams(&before);
    let b = streams(&after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, data) in a {
        if name != ["\u{5}DocumentSummaryInformation"] {
            assert!(
                b.get(&name) == Some(&data),
                "signature edit changed {}",
                name.join("/")
            );
        }
    }
    let a = p.property_sets().unwrap();
    let b = q.property_sets().unwrap();
    let without_signature = |p: &litchi_doc::package::property_set::Snapshot| {
        let mut s = p
            .document_summary_information()
            .unwrap()
            .unwrap()
            .section()
            .clone();
        s.remove(DIGITAL_SIGNATURE);
        s
    };
    assert_eq!(without_signature(&a), without_signature(&b));
    assert_eq!(
        a.user_defined_properties().unwrap(),
        b.user_defined_properties().unwrap()
    );
    assert_eq!(
        p.apply_property_sets(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        after
    );
    assert_eq!(
        q.apply_property_sets(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        before
    );
    q
}

#[test]
fn native_doc_vba_signature_property_storage_crud_keeps_opaque_bytes_and_macro_streams() {
    use litchi_doc::vba_signature as sig;
    for file in MACROS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(
            p.property_sets()
                .unwrap()
                .vba_signature()
                .unwrap()
                .is_none()
        );
        let value = sig::Snapshot::parse_property(&signature_blob(false)).unwrap();
        let mut e = p.edit_property_sets().unwrap();
        assert!(e.set_vba_signature(&value).unwrap());
        let c = e.commit().unwrap();
        let added = properties_publish(&p, &c);
        assert_eq!(
            added
                .property_sets()
                .unwrap()
                .vba_signature()
                .unwrap()
                .unwrap(),
            value
        );
        let mut e = added.edit_property_sets().unwrap();
        assert!(!e.set_vba_signature(&value).unwrap());
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = added.edit_property_sets().unwrap();
        assert!(
            e.edit_vba_signature(sig::Limits::default(), |v| v
                .set_signature([0x76, 0x77, 0x78, 0x79])
                .map(|_| ()))
                .unwrap()
        );
        let staged = e.vba_signature(sig::Limits::default()).unwrap().unwrap();
        assert!(
            matches!(e.edit_vba_signature(sig::Limits::default(),|v|{v.set_certificate_store([0x99])?;Err(sig::Error::Invalid("intentional callback failure".into()))}),Err(litchi_doc::package::property_set::Error::Package(litchi_doc::Error::InvalidFormat(m))) if m=="invalid VBA signature edit: invalid VBA signature blob: intentional callback failure")
        );
        assert_eq!(
            e.vba_signature(sig::Limits::default()).unwrap().unwrap(),
            staged
        );
        let c = e.commit().unwrap();
        let updated = properties_publish(&added, &c);
        let current = updated
            .property_sets()
            .unwrap()
            .vba_signature()
            .unwrap()
            .unwrap();
        assert_eq!(current.info().signature(), [0x76, 0x77, 0x78, 0x79]);
        assert_eq!(current.info().certificate_store(), [4, 5, 6, 7]);
        assert_eq!(current.info().reserved_timestamp_marker(), 0xa5a5_5a5a);
        assert!(current.bytes().contains(&0xcc));
        assert!(matches!(
            updated.apply_property_sets(c.patch()),
            Err(litchi_doc::package::property_set::Error::PatchConflict)
        ));
        let mut e = updated.edit_property_sets().unwrap();
        assert!(e.remove_vba_signature().unwrap());
        assert!(!e.remove_vba_signature().unwrap());
        let c = e.commit().unwrap();
        let removed = properties_publish(&updated, &c);
        assert!(
            removed
                .property_sets()
                .unwrap()
                .vba_signature()
                .unwrap()
                .is_none()
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

fn xst(s: &str) -> Vec<u8> {
    let units = s.encode_utf16().collect::<Vec<_>>();
    let mut result = u16::try_from(units.len()).unwrap().to_le_bytes().to_vec();
    result.extend(units.into_iter().flat_map(u16::to_le_bytes));
    result
}
fn stw_user(duplicate: bool) -> Vec<u8> {
    let mut result = vec![0xff, 0xff, 4, 0, 4, 0];
    for name in [
        "Sign",
        if duplicate { "Sign" } else { "SigAgile" },
        "SigV3",
        "Other",
    ] {
        result.extend(xst(name));
        result.extend([0xa5, 0x5a, 0xc3, 0x3c]);
    }
    for _ in 0..3 {
        result.extend(signature_blob(true));
    }
    result.extend(xst("unrecognized value"));
    result
}
fn seed_stw_user(p: &Package, data: Vec<u8>) -> Package {
    use litchi_ole_common::object;
    let bytes = p.to_bytes().unwrap();
    let all = streams(&bytes);
    let table = table_name(p);
    let mut word = all[&vec!["WordDocument".into()]].clone();
    let mut table_bytes = all[&table].clone();
    let offset = u32::try_from(table_bytes.len()).unwrap();
    let length = u32::try_from(data.len()).unwrap();
    table_bytes.extend(data);
    let pair = 154 + litchi_doc::parts::vba_signature::FIB_INDEX_STW_USER * 8;
    word[pair..pair + 4].copy_from_slice(&offset.to_le_bytes());
    word[pair + 4..pair + 8].copy_from_slice(&length.to_le_bytes());
    let original = object::Package::open(bytes.clone()).unwrap();
    let mut e = original.edit().unwrap();
    e.put_stream(&["WordDocument".into()], word).unwrap();
    e.put_stream(&table, table_bytes).unwrap();
    let c = e.commit().unwrap();
    let q = reopen(&Package::from_bytes(c.patch().after()).unwrap());
    preserved(p, &q, litchi_doc::parts::vba_signature::FIB_INDEX_STW_USER);
    let after = object::Package::open(q.to_bytes().unwrap()).unwrap();
    assert_eq!(
        c.patch().inverse().apply(&after).unwrap().patch().after(),
        bytes
    );
    q
}

#[test]
fn native_doc_stw_user_signature_names_keep_padding_and_defer_invalid_metadata() {
    use litchi_doc::SignatureName;
    for file in MACROS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(p.document().vba_signatures().unwrap().is_none());
        let data = stw_user(false);
        let q = seed_stw_user(&p, data.clone());
        let document = q.document();
        let signatures = document.vba_signatures().unwrap().unwrap();
        assert_eq!(signatures.len(), 3);
        assert_eq!(
            payload(&q, litchi_doc::parts::vba_signature::FIB_INDEX_STW_USER),
            data
        );
        for name in [
            SignatureName::Sign,
            SignatureName::SigAgile,
            SignatureName::SigV3,
        ] {
            let value = signatures.get(name).unwrap();
            assert_eq!(value.name_str(), name.as_str());
            assert_eq!(value.bytes(), signature_blob(true));
            assert_eq!(value.snapshot().info().signature(), [9, 8]);
            assert_eq!(value.snapshot().info().certificate_store(), [7, 6, 5]);
            assert_eq!(value.snapshot().info().padding(), [0xee]);
        }
        assert!(
            q.property_sets()
                .unwrap()
                .vba_signature()
                .unwrap()
                .is_none()
        );
        let bad = seed_stw_user(&p, stw_user(true));
        assert_eq!(bad.document().text().unwrap(), p.document().text().unwrap());
        for _ in 0..2 {
            assert!(
                matches!(bad.document().vba_signatures(),Err(litchi_doc::Error::Corrupted(m)) if m=="invalid VBA signature metadata: Corrupted file: StwUser name 1 is not unique")
            );
        }
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_binary_word_corpus_has_excel_ole_objects_but_no_activex_control_specimen() {
    use litchi_doc::{OpenOptions, Password, embedded_object::Metadata};
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office");
    let manifest: serde_json::Value =
        serde_json::from_slice(&std::fs::read(root.join("manifest.json")).unwrap()).unwrap();
    let mut count = 0;
    let mut objects = 0;
    let mut encrypted = 0;
    for item in manifest["files"].as_array().unwrap() {
        let name = item["path"].as_str().unwrap();
        if !matches!(
            Path::new(name).extension().and_then(|x| x.to_str()),
            Some("doc" | "dot")
        ) {
            continue;
        }
        let bytes = std::fs::read(root.join(name)).unwrap();
        count += 1;
        let p = if name == "word/word-encrypted.doc" {
            encrypted += 1;
            Package::from_bytes_with_options(
                &bytes,
                OpenOptions::default().password(Password::new("Office42!".into())),
            )
            .unwrap()
        } else {
            Package::from_bytes(&bytes).unwrap()
        };
        if name == "word/word-encrypted.doc" {
            assert!(
                !cfb_metadata::directory(&bytes)
                    .keys()
                    .any(|p| p.first().is_some_and(|s| s == "ObjectPool"))
            );
        } else {
            let snapshot = p.embedded_objects().unwrap();
            let raw_storages = cfb_metadata::directory(&bytes)
                .keys()
                .filter(|path| path.len() == 2 && path[0] == "ObjectPool")
                .map(|path| path[1].clone())
                .collect::<std::collections::BTreeSet<_>>();
            assert_eq!(
                snapshot
                    .storages()
                    .iter()
                    .map(|object| object.key().to_owned())
                    .collect::<std::collections::BTreeSet<_>>(),
                raw_storages
            );
            for object in snapshot.storages().iter() {
                objects += 1;
                assert_eq!(name, "word/word-visual-objects.doc");
                let metadata = Metadata::of(object).unwrap();
                assert!(!metadata.is_activex());
                let info = metadata
                    .obj_info()
                    .expect("native ObjInfo must decode, not fall back to opaque bytes");
                assert!(!info.activex && !info.stream_control);
                assert!(matches!(
                    metadata.class_id(),
                    Some(
                        "00020830-0000-0000-C000-000000000046"
                            | "00020821-0000-0000-C000-000000000046"
                    )
                ));
            }
        }
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(std::fs::read(root.join(name)).unwrap(), bytes);
    }
    assert_eq!((count, objects, encrypted), (20, 2, 1));
}
