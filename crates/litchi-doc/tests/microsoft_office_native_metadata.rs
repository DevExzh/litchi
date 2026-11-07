//! Native binary Word properties and inert table metadata.
use litchi_doc::parts::{
    associated_strings::{AssociatedStringSlot, DocumentAssociatedStrings},
    dofr::{DofrPayload, DofrType},
};
use litchi_doc::user_defined_hyperlinks::{
    Hyperlink, Limits, MutationError, UserDefinedHyperlinks,
};
use litchi_doc::{FieldStory, FieldType, HyperlinkAssociation, Package};
use litchi_ole_common::property_set::{
    CodePage, Value,
    user_defined::{self, Properties},
};
use std::{collections::BTreeMap, io::Cursor};
const FILES: [&str; 6] = [
    "word-text-structure.doc",
    "word-math-navigation.doc",
    "word-review-controls.doc",
    "word-recorded-macro.doc",
    "word-recorded-macro.dot",
    "word-layout-extensions.dot",
];
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
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
            let n = p.iter().map(String::as_str).collect::<Vec<_>>();
            let b = ole.open_stream(&n).unwrap();
            (p, b)
        })
        .collect()
}
fn table(p: &Package, index: usize) -> Vec<u8> {
    let d = p.document();
    let (at, len) = d.fib().get_table_pointer(index).unwrap();
    let s = streams(&p.to_bytes().unwrap());
    let key = vec![format!("{}Table", u8::from(d.fib().which_table_stream()))];
    s[&key][at as usize..(at + len) as usize].to_vec()
}
fn assert_non_properties(before: &[u8], after: &[u8]) {
    let a = streams(before);
    let b = streams(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (path, data) in a {
        if !matches!(path.as_slice(),[n] if n=="\u{5}SummaryInformation" || n=="\u{5}DocumentSummaryInformation")
        {
            assert!(data == b[&path], "changed {}", path.join("/"));
        }
    }
}
fn save(p: &Package) -> Package {
    let f = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    p.save(f.path()).unwrap();
    Package::open(f.path()).unwrap()
}
fn links(p: &Package) -> UserDefinedHyperlinks {
    p.property_sets()
        .unwrap()
        .user_defined_hyperlinks(p.document().fields_table(), Limits::default())
        .unwrap()
        .unwrap()
}
fn blob(p: &Package) -> Vec<u8> {
    let s = p
        .property_sets()
        .unwrap()
        .user_defined_properties()
        .unwrap()
        .unwrap();
    let Value::Blob(b) = s.find_named(user_defined::HYPERLINKS).unwrap().1 else {
        panic!("not a BLOB")
    };
    b.clone()
}

#[test]
fn native_reserved_hyperlinks_decode_the_single_ole_blob_envelope() {
    for (name, expected, raw_len) in [
        (
            FILES[0],
            vec![(
                0,
                7_012_461,
                "https://example.com/litchi-fixture",
                "",
                317,
                365,
                380,
            )],
            128,
        ),
        (
            FILES[1],
            vec![
                (8, 1_769_520, "", "_Toc242967667", 233, 265, 330),
                (2, 1_769_520, "", "_Toc242967666", 139, 171, 231),
            ],
            164,
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let links = links(&p);
        assert_eq!(links.entries().len(), expected.len());
        let raw = blob(&p);
        assert_eq!(raw.len(), raw_len);
        assert_eq!(
            u32::from_le_bytes(raw[..4].try_into().unwrap()),
            u32::try_from(expected.len() * 6).unwrap()
        );
        for (link, (app, hash, target, location, start, separator, end)) in
            links.entries().iter().zip(expected)
        {
            assert_eq!(link.app(), app);
            assert_eq!(link.stored_hash(), hash);
            assert_eq!(link.office_art(), 0);
            assert_eq!(link.info(), 5);
            assert_eq!(link.target(), target);
            assert_eq!(link.location(), location);
            let HyperlinkAssociation::FieldCandidates(c) = link.association() else {
                panic!("native field candidate")
            };
            assert_eq!(c.len(), 1);
            let c = &c[0];
            assert_eq!(c.story, FieldStory::Main);
            assert_eq!(c.plcfld_index, app as u32);
            assert_eq!(c.begin_cp, start);
            assert_eq!(c.field.field_type, FieldType::Hyperlink);
            assert_eq!(c.field.separator_cp, Some(separator));
            assert_eq!(c.field.end_cp, end);
            let field = p.document().field_text(&c.field).unwrap();
            assert!(
                field
                    .instruction
                    .contains(if target.is_empty() { location } else { target })
            );
        }
        let mut section = p
            .property_sets()
            .unwrap()
            .user_defined_properties()
            .unwrap()
            .unwrap();
        let shared = Properties::new(&section, Limits::default())
            .unwrap()
            .hyperlinks()
            .unwrap()
            .unwrap();
        user_defined::Edit::new(&mut section, Limits::default())
            .unwrap()
            .set_hyperlinks(shared)
            .unwrap();
        assert_eq!(
            section.find_named(user_defined::HYPERLINKS).unwrap().1,
            &Value::Blob(raw)
        );
        let mut noop = p.edit_property_sets().unwrap();
        assert!(
            !noop
                .put_user_defined_hyperlinks(&links, Limits::default())
                .unwrap()
        );
        assert_eq!(noop.commit().unwrap().package().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_hyperlink_property_edits_require_resolved_fields_and_preserve_body_streams() {
    for name in &FILES[..2] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let original = links(&p);
        let mut updated = original.clone();
        let first = &original.entries()[0];
        let replacement = Hyperlink::new(
            first.app(),
            first.office_art(),
            first.info(),
            "https://example.test/native-property-edit",
            first.location(),
        )
        .unwrap();
        updated
            .replace_hyperlink(0, replacement, p.document().fields_table())
            .unwrap();
        let mut e = p.edit_property_sets().unwrap();
        e.edit_summary_information(|s| s.set_title("Earlier staged metadata"))
            .unwrap();
        assert!(matches!(
            e.put_user_defined_hyperlinks(&updated, Limits::default()),
            Err(MutationError::UnresolvedFieldCandidates)
        ));
        for entry in updated.entries_mut() {
            let HyperlinkAssociation::FieldCandidates(c) = entry.association() else {
                panic!("field candidate")
            };
            assert_eq!(c.len(), 1);
            let c = c[0].clone();
            // The read test independently ties these native candidates to field text.
            entry.resolve_field(c.story, c.plcfld_index).unwrap();
        }
        assert!(
            e.put_user_defined_hyperlinks(&updated, Limits::default())
                .unwrap()
        );
        let commit = e.commit().unwrap();
        let after = save(commit.package());
        let output = after.to_bytes().unwrap();
        assert_eq!(
            after.metadata().title.as_deref(),
            Some("Earlier staged metadata")
        );
        let reread = links(&after);
        assert_eq!(reread.entries().len(), original.entries().len());
        assert_eq!(
            reread.entries()[0].target(),
            "https://example.test/native-property-edit"
        );
        assert!(reread.entries()[0].hash_matches());
        assert_eq!(
            reread.entries().iter().map(|e| e.app()).collect::<Vec<_>>(),
            original
                .entries()
                .iter()
                .map(|e| e.app())
                .collect::<Vec<_>>()
        );
        assert_non_properties(&bytes, &output);
        assert_eq!(
            p.apply_property_sets(commit.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            output
        );
        assert!(matches!(
            after.apply_property_sets(commit.patch()),
            Err(litchi_doc::package::property_set::Error::PatchConflict)
        ));
        assert_eq!(
            after
                .apply_property_sets(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        let mut e = after.edit_property_sets().unwrap();
        assert!(e.remove_user_defined_hyperlinks().unwrap());
        assert!(!e.remove_user_defined_hyperlinks().unwrap());
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert!(
            absent
                .property_sets()
                .unwrap()
                .user_defined_hyperlinks(absent.document().fields_table(), Limits::default())
                .unwrap()
                .is_none()
        );
        assert_non_properties(&output, &absent.to_bytes().unwrap());
        assert_eq!(
            absent
                .apply_property_sets(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            output
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_summary_properties_keep_mac_codepage_dates_templates_and_saved_counts() {
    for (name, words, chars, date, template) in [
        (FILES[0], 59, 340, "2026-10-15T05:23:00Z", "Normal.dotm"),
        (
            FILES[1],
            46,
            607,
            "2026-10-15T06:43:00Z",
            "word-math-navigation.dotx",
        ),
        (FILES[2], 18, 239, "2026-10-15T05:53:00Z", "Normal.dotm"),
        (
            FILES[3],
            12,
            169,
            "2026-10-15T06:00:00Z",
            "word-recorded-macro.dotm",
        ),
        (FILES[4], 12, 169, "2026-10-15T06:01:00Z", "Normal.dotm"),
        (
            FILES[5],
            27,
            352,
            "2026-10-15T11:26:00Z",
            "word-layout-extensions.dot",
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let m = p.metadata();
        let review = name == FILES[2];
        assert_eq!(
            m.author.as_deref(),
            Some(if review {
                "Litchi Fixture Author"
            } else {
                "Ryker Zhu"
            })
        );
        assert_eq!(m.last_modified_by.as_deref(), Some("Ryker Zhu"));
        assert_eq!(m.application.as_deref(), Some("Microsoft Office Word"));
        assert_eq!(m.codepage, Some(10000));
        assert_eq!(m.revision.as_deref(), Some("2"));
        assert_eq!(m.template.as_deref(), Some(template));
        assert_eq!(m.created.as_ref().unwrap().to_string(), date);
        assert_eq!(m.modified.as_ref().unwrap().to_string(), date);
        assert_eq!(m.page_count, Some(1));
        assert_eq!(m.word_count, Some(words));
        assert_eq!(m.character_count, Some(chars));
        assert_eq!(m.security, Some(0));
        if review {
            assert_eq!(m.title.as_deref(), Some("Native Word review fixture"));
            assert_eq!(m.keywords, ["native", "word", "review", "003"]);
            assert_eq!(m.company.as_deref(), Some("Litchi Native Fixtures"));
        }
        let mut e = p.edit_metadata().unwrap();
        e.properties_mut()
            .unwrap()
            .set("Title", m.title.as_deref().unwrap())
            .unwrap();
        assert_eq!(e.commit().unwrap().package().to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_metadata_crud_preserves_reserved_hyperlinks_macros_and_staged_changes() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let original = p.metadata().clone();
        let old_section = p
            .property_sets()
            .unwrap()
            .user_defined_properties()
            .unwrap();
        let old_hlinks = old_section
            .as_ref()
            .and_then(|s| s.find_named(user_defined::HYPERLINKS))
            .map(|(_, v)| v.clone());
        let mut e = p.edit_metadata().unwrap();
        let props = e.properties_mut().unwrap();
        props.set("Title", "Native property review").unwrap();
        props.set("Company", "Fixture review").unwrap();
        props
            .set_user_defined("NativeReviewed", Value::Bool(true))
            .unwrap();
        props
            .set_user_defined("NativeUnicode", Value::Lpwstr("日本語".into()))
            .unwrap();
        assert!(matches!(
            e.merge_metadata(&litchi_core::Metadata {
                identifier: Some("not-storable".into()),
                ..Default::default()
            }),
            Err(litchi_doc::package::metadata::Error::Metadata(
                litchi_ole_common::property_set::summary_properties::MetadataError::Unsupported {
                    field: "identifier"
                }
            ))
        ));
        let c = e.commit().unwrap();
        let after = save(c.package());
        let output = after.to_bytes().unwrap();
        assert_non_properties(&bytes, &output);
        assert_eq!(
            after.metadata().title.as_deref(),
            Some("Native property review")
        );
        assert_eq!(after.metadata().company.as_deref(), Some("Fixture review"));
        assert_eq!(after.metadata().author, original.author);
        assert_eq!(after.metadata().created, original.created);
        assert_eq!(after.metadata().codepage, Some(10000));
        let section = after
            .property_sets()
            .unwrap()
            .user_defined_properties()
            .unwrap()
            .unwrap();
        assert_eq!(section.page(), Some(CodePage::Utf16Le));
        assert_eq!(
            section.find_named("nativereviewed").unwrap().1,
            &Value::Bool(true)
        );
        assert_eq!(
            section.find_named("nativeunicode").unwrap().1,
            &Value::Lpwstr("日本語".into())
        );
        assert_eq!(
            section.find_named(user_defined::HYPERLINKS).map(|(_, v)| v),
            old_hlinks.as_ref()
        );
        assert_eq!(
            p.apply_metadata(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            output
        );
        assert!(matches!(
            after.apply_metadata(c.patch()),
            Err(litchi_doc::package::metadata::Error::PatchConflict)
        ));
        assert_eq!(
            after
                .apply_metadata(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        let mut e = after.edit_metadata().unwrap();
        let props = e.properties_mut().unwrap();
        assert!(props.remove("Title").unwrap());
        assert!(props.remove("NativeReviewed").unwrap());
        assert!(props.remove("NativeUnicode").unwrap());
        let removed = e.commit().unwrap();
        let absent = save(removed.package());
        assert_eq!(absent.metadata().title, None);
        assert_non_properties(&output, &absent.to_bytes().unwrap());
        assert_eq!(
            absent
                .apply_metadata(&removed.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            output
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_annotation_tags_and_list_style_records_keep_exact_table_bytes() {
    for (name, ids, has_list) in [
        (FILES[0], vec![1_418_134_056], true),
        (FILES[1], vec![], true),
        (FILES[2], vec![1_194_253_998, 2_055_799_712], false),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.document();
        let actual = d
            .annotation_bookmarks()
            .unwrap()
            .map(|t| t.entries().iter().map(|t| t.id().raw()).collect::<Vec<_>>())
            .unwrap_or_default();
        assert_eq!(actual, ids);
        if !ids.is_empty() {
            let tags = d.annotation_bookmarks().unwrap().unwrap();
            assert_eq!(
                tags.to_bytes(&litchi_doc::SourceLimits::default()).unwrap(),
                table(&p, 37)
            );
        }
        if has_list {
            let records = d.dofr_records().unwrap().unwrap();
            assert_eq!(records.len(), 1);
            assert_eq!(records.bytes(), table(&p, 99));
            let record = records.get(0).unwrap();
            assert_eq!(record.cb(), 16);
            assert_eq!(record.kind(), DofrType::ListStyles);
            let DofrPayload::ListStyles(styles) = record.payload().unwrap() else {
                panic!("list styles")
            };
            assert_eq!(styles.len(), 1);
            let style = styles.get(0).unwrap();
            assert_eq!(style.ilst(), 0);
            assert_eq!(style.istd_list(), 12);
            assert!(!style.style_defined());
            assert_eq!(style.unused(), 0);
        } else {
            assert!(d.dofr_records().unwrap().is_none());
        }
        let mut e = p.edit_metadata().unwrap();
        e.properties_mut()
            .unwrap()
            .set("Title", "Native table metadata preservation")
            .unwrap();
        let c = e.commit().unwrap();
        let after = save(c.package());
        assert_non_properties(&bytes, &after.to_bytes().unwrap());
        assert_eq!(
            after.document().annotation_bookmarks().unwrap(),
            d.annotation_bookmarks().unwrap()
        );
        assert_eq!(table(&after, 99), table(&p, 99));
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_associated_string_tables_round_trip_without_conflating_summary_templates() {
    for name in FILES {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let document = p.document();
        let strings = document.associated_strings().unwrap();
        let review = name == FILES[2];
        assert_eq!(
            strings.author(),
            if review {
                "Litchi Fixture Author"
            } else {
                "Ryker Zhu"
            }
        );
        assert_eq!(strings.last_revised_by(), "Ryker Zhu");
        assert_eq!(
            strings.title(),
            if review {
                "Native Word review fixture"
            } else {
                ""
            }
        );
        assert_eq!(
            strings.subject(),
            if review {
                "Tracked revisions, bookmarks, fields, comments"
            } else {
                ""
            }
        );
        assert_eq!(
            strings.keywords(),
            if review { "native;word;review;003" } else { "" }
        );
        assert_eq!(strings.template_path(), "");
        assert!(!p.metadata().template.as_ref().unwrap().is_empty());
        assert_eq!(strings.iter().len(), 18);
        assert_eq!(strings.to_bytes().unwrap(), table(&p, 32));
        assert_eq!(
            &DocumentAssociatedStrings::parse_bytes(&strings.to_bytes().unwrap()).unwrap(),
            strings
        );
        let mut detached = strings.clone();
        detached
            .set(AssociatedStringSlot::Title, "Detached title")
            .unwrap();
        assert_eq!(
            DocumentAssociatedStrings::parse_bytes(&detached.to_bytes().unwrap())
                .unwrap()
                .title(),
            "Detached title"
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}
