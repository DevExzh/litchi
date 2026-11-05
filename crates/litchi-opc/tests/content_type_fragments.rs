//! Durable content-type deltas retain source spelling and destination overlays.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use litchi_opc::{
    ContentTypeFragmentEdit as Edit, ContentTypeFragmentKey as Key, ContentTypeFragments,
    OpcPackage, PackageWriter, SourceLimitKind, SourceLimits,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::sync::Arc;
const ROOT: &str = "<t:Types xmlns:t='http://schemas.openxmlformats.org/package/2006/content-types' xmlns:x='urn:extra'>";
const BASE: &str = "<t:Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><t:Default Extension='bin' ContentType='application/octet-stream'/>";
const A: &str = "<t:Default ContentType='application/x-one' Extension='X' x:keep='yes'/>";
const B: &str = "<t:Override PartName=\"/custom/ITEM.bin\" ContentType='application/octet-stream'></t:Override>";
const C: &str = "<t:Default Extension='Y' ContentType='application/x-two'/>";
fn xml(body: &str) -> String {
    format!("\u{feff}{ROOT}{BASE}<?head retained?>{body}</t:Types><!--end-->")
}
fn catalog(xml: &str) -> ContentTypeFragments {
    ContentTypeFragments::from_bytes(Arc::new(xml.as_bytes().to_vec()), SourceLimits::default())
        .unwrap()
}
fn package(xml: &str) -> OpcPackage {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_deflated("[Content_Types].xml", xml.as_bytes())
        .unwrap();
    writer.write_deflated("custom/item.bin", b"opaque").unwrap();
    writer.write_deflated("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#).unwrap();
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), SourceLimits::default()).unwrap()
}
#[test]
fn publication_inverse_preserves_destination_changes_and_exact_mapping_bytes() {
    let original = xml(&format!("{A}\n<!--before-->{B}<x:kept/>{C}\n"));
    let mut package = package(&original);
    let source = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let fragments = source.fragments(limits).unwrap();
    assert_eq!(fragments.iter().len(), 5);
    let b = fragments.get(Key::Override("/CUSTOM/item.BIN")).unwrap();
    assert_eq!(b.unwrap().bytes(), B.as_bytes());
    assert_eq!(b.unwrap().key(), Key::Override("/custom/item.bin"));
    assert_eq!(b.unwrap().following_key(), Some(Key::Default("y")));
    assert_eq!(b.unwrap().trailing_bytes(), b"<x:kept/>");
    let removed = source
        .plan_fragment_edits(
            &[Edit {
                before: b,
                after: None,
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    package
        .try_replace_content_types(source.bytes(), &removed, SourceLimits::default())
        .unwrap();
    let saved = PackageWriter::to_bytes(&package).unwrap();
    let mut reopened = OpcPackage::from_bytes(&saved, SourceLimits::default()).unwrap();
    let current = reopened
        .source_content_types(SourceLimits::default())
        .unwrap();
    let overlay = original
        .replace(B, "")
        .replace("head retained", "head overlay")
        .replace("application/x-one", "application/x-changed");
    reopened
        .try_replace_content_types_bytes(current.bytes(), Arc::new(overlay.into_bytes()), limits)
        .unwrap();
    let current = reopened
        .source_content_types(SourceLimits::default())
        .unwrap();
    let restored = current
        .plan_fragment_edits(
            &[Edit {
                before: None,
                after: b,
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    let expected = original
        .replace("head retained", "head overlay")
        .replace("application/x-one", "application/x-changed");
    assert_eq!(restored.bytes(), expected.as_bytes());
    reopened
        .try_replace_content_types(current.bytes(), &restored, SourceLimits::default())
        .unwrap();
    let saved = PackageWriter::to_bytes(&reopened).unwrap();
    assert_eq!(
        ArchiveReader::new(&saved, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .read("[Content_Types].xml")
            .unwrap(),
        expected.as_bytes()
    );
    assert_eq!(
        OpcPackage::from_bytes(&saved, SourceLimits::default())
            .unwrap()
            .source_content_types(SourceLimits::default())
            .unwrap()
            .bytes(),
        expected.as_bytes()
    );
}
#[test]
fn all_deletion_subsets_restore_in_any_batch_order_ahead_of_unrelated_append() {
    let original = xml(&format!("{A}\n{B}<!--gap-->{C}\n"));
    let fragments = catalog(&original);
    let source = package(&original)
        .source_content_types(SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let append = "<t:Default Extension='z' ContentType='application/x-extra'/>";
    for mask in 1..8 {
        let selected: Vec<_> = [
            Key::Default("x"),
            Key::Override("/custom/item.bin"),
            Key::Default("y"),
        ]
        .into_iter()
        .enumerate()
        .filter(|(index, _)| mask & (1 << index) != 0)
        .map(|(_, key)| fragments.get(key).unwrap())
        .collect();
        let edits: Vec<_> = selected
            .iter()
            .map(|&before| Edit {
                before,
                after: None,
            })
            .collect();
        let removed = source
            .plan_fragment_edits(&edits, limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        let overlay = String::from_utf8(removed.bytes().to_vec())
            .unwrap()
            .replace("</t:Types>", &format!("{append}</t:Types>"));
        let current = package(&overlay)
            .source_content_types(SourceLimits::default())
            .unwrap();
        let mut edits: Vec<_> = selected
            .iter()
            .map(|&after| Edit {
                before: None,
                after,
            })
            .collect();
        for _ in 0..2 {
            edits.reverse();
            let restored = current
                .plan_fragment_edits(&edits, limits)
                .unwrap()
                .materialize(limits)
                .unwrap();
            assert_eq!(
                restored.bytes(),
                original
                    .replace("</t:Types>", &format!("{append}</t:Types>"))
                    .as_bytes(),
                "mask {mask}"
            );
        }
    }
}
#[test]
fn native_key_normalization_and_duplicate_identity_guards_agree() {
    let source = xml(&format!("{A}{B}"));
    let fragments = catalog(&source);
    assert_eq!(
        fragments.get(Key::Default("X")).unwrap().unwrap().key(),
        Key::Default("x")
    );
    assert!(fragments.get(Key::Override("x")).unwrap().is_none());
    let spaced = catalog(&xml(&A.replace("Extension='X'", "Extension=' X '")));
    assert!(spaced.get(Key::Default("X")).unwrap().is_none());
    assert!(spaced.get(Key::Default(" X ")).unwrap().is_some());
    let escaped = catalog(&xml(&A.replace("Extension='X'", "Extension='&#x58;'")));
    assert!(escaped.get(Key::Default("x")).unwrap().is_some());
    for duplicate in [
        xml(&format!(
            "{A}{}",
            A.replace("Extension='X'", "Extension='x'")
        )),
        xml(&format!(
            "{B}{}",
            B.replace("/custom/ITEM.bin", "/CUSTOM/item.BIN")
        )),
    ] {
        assert!(
            ContentTypeFragments::from_bytes(
                Arc::new(duplicate.into_bytes()),
                SourceLimits::default()
            )
            .is_err()
        );
    }
}
#[test]
fn simultaneous_replacements_keep_destination_slots_and_unrelated_gaps() {
    let original = xml(&format!("{A}\n{B}<!--gap-->{C}"));
    let changed = original
        .replace("application/x-", "application/new-")
        .replace("Extension='X'", "Extension='x'");
    let before = catalog(&original);
    let after = catalog(&changed);
    let destination = original.replace("<!--gap-->", "<?overlay kept?>");
    let source = package(&destination)
        .source_content_types(SourceLimits::default())
        .unwrap();
    let edits: Vec<_> = [
        Key::Default("y"),
        Key::Default("X"),
        Key::Override("/custom/item.bin"),
    ]
    .into_iter()
    .map(|key| Edit {
        before: before.get(key).unwrap(),
        after: after.get(key).unwrap(),
    })
    .collect();
    let result = source
        .plan_fragment_edits(&edits, SourceLimits::default())
        .unwrap()
        .materialize(SourceLimits::default())
        .unwrap();
    assert_eq!(
        result.bytes(),
        changed.replace("<!--gap-->", "<?overlay kept?>").as_bytes()
    );
}
#[test]
fn exact_noops_share_bytes_but_recheck_tightened_xml_limits() {
    let original = xml(A);
    let fragments = catalog(&original);
    let source = package(&original)
        .source_content_types(SourceLimits::default())
        .unwrap();
    let item = fragments.get(Key::Default("x")).unwrap();
    let limits = SourceLimits::default();
    let plan = source
        .plan_fragment_edits(
            &[Edit {
                before: item,
                after: item,
            }],
            limits,
        )
        .unwrap();
    assert!(plan.is_noop());
    let result = plan.materialize(limits).unwrap();
    assert!(std::ptr::eq(result.bytes(), source.bytes()));
    let shallow = SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        source
            .plan_fragment_edits(&[], limits)
            .unwrap()
            .materialize(shallow)
            .is_err()
    );
    assert!(
        source
            .plan_edit(&[], &[], limits)
            .unwrap()
            .materialize(shallow)
            .is_err()
    );
}
#[test]
fn output_byte_mapping_event_and_entry_limits_are_admitted_before_materialization() {
    let original = xml(A);
    let expanded = xml(&format!("{A}{B}{C}"));
    let after = catalog(&expanded);
    let source = package(&original)
        .source_content_types(SourceLimits::default())
        .unwrap();
    let edits = [
        Edit {
            before: None,
            after: after.get(Key::Override("/custom/item.bin")).unwrap(),
        },
        Edit {
            before: None,
            after: after.get(Key::Default("y")).unwrap(),
        },
    ];
    let plan = source
        .plan_fragment_edits(&edits, SourceLimits::default())
        .unwrap();
    assert_eq!(plan.final_len(), expanded.len());
    assert_eq!(plan.mapping_count(), 5);
    let events = plan.event_upper_bound();
    let exact = SourceLimits::builder()
        .max_content_types_bytes(expanded.len())
        .unwrap()
        .max_content_type_mappings(5)
        .unwrap()
        .max_events(events)
        .unwrap()
        .build()
        .unwrap();
    assert_eq!(
        source
            .plan_fragment_edits(&edits, exact)
            .unwrap()
            .materialize(exact)
            .unwrap()
            .bytes(),
        expanded.as_bytes()
    );
    for (limits, kind) in [
        (
            SourceLimits::builder()
                .max_content_types_bytes(expanded.len() - 1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::ContentTypesBytes,
        ),
        (
            SourceLimits::builder()
                .max_content_type_mappings(4)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::ContentTypeMappings,
        ),
        (
            SourceLimits::builder()
                .max_events(events - 1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::Events,
        ),
        (
            SourceLimits::builder()
                .max_entry_bytes(expanded.len() as u64 - 1)
                .unwrap()
                .build()
                .unwrap(),
            SourceLimitKind::EntryBytes,
        ),
    ] {
        assert!(
            matches!(source.plan_fragment_edits(&edits, limits), Err(litchi_opc::OpcError::Limit {kind: observed, ..}) if observed == kind)
        );
        assert!(
            source
                .plan_fragment_edits(&edits, SourceLimits::default())
                .unwrap()
                .materialize(limits)
                .is_err()
        );
    }
}
#[test]
fn stale_entries_context_keys_and_terminal_gaps_refuse_without_mutation() {
    let original = xml(&format!("{A}<!--gap-->{B}"));
    let before = catalog(&original);
    let source = package(&original)
        .source_content_types(SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let key = Key::Default("x");
    let a = before.get(key).unwrap();
    for stale in [
        original.replace("application/x-one", "application/stale"),
        original.replace("urn:extra", "urn:other"),
        original.replace("<t:Types ", "<t:Types xml:space='preserve' "),
    ] {
        let stale = catalog(&stale);
        assert!(
            source
                .plan_fragment_edits(
                    &[Edit {
                        before: stale.get(key).unwrap(),
                        after: None
                    }],
                    limits
                )
                .is_err()
        );
    }
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: a,
                    after: before.get(Key::Override("/custom/item.bin")).unwrap()
                }],
                limits
            )
            .is_err()
    );
    assert!(
        source
            .plan_fragment_edits(
                &[
                    Edit {
                        before: a,
                        after: None
                    },
                    Edit {
                        before: a,
                        after: None
                    }
                ],
                limits
            )
            .is_err()
    );
    let destination = xml(&format!("{A}<!--changed-->{C}"));
    let destination_source = package(&destination)
        .source_content_types(SourceLimits::default())
        .unwrap();
    assert!(
        destination_source
            .plan_fragment_edits(
                &[Edit {
                    before: None,
                    after: before.get(Key::Override("/custom/item.bin")).unwrap()
                }],
                limits
            )
            .is_err()
    );
    assert_eq!(source.bytes(), original.as_bytes());
}
#[test]
fn empty_aliased_root_expands_without_canonicalizing_its_wrapper() {
    let original = format!("\u{feff}{} /><!--end-->", ROOT.trim_end_matches('>'));
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_deflated("[Content_Types].xml", original.as_bytes())
        .unwrap();
    let package =
        OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), SourceLimits::default()).unwrap();
    let source = package
        .source_content_types(SourceLimits::default())
        .unwrap();
    let after = catalog(&format!("{ROOT}{A}</t:Types>"));
    let limits = SourceLimits::default();
    let result = source
        .plan_fragment_edits(
            &[Edit {
                before: None,
                after: after.get(Key::Default("x")).unwrap(),
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    assert_eq!(
        result.bytes(),
        format!(
            "\u{feff}{} >{A}</t:Types><!--end-->",
            ROOT.trim_end_matches('>')
        )
        .as_bytes()
    );
}
#[test]
fn malformed_or_unbound_catalogs_never_grant_edit_authority() {
    for source in [
        xml(A).replace("Extension='X'", ""),
        xml(A).replace("t:Default ", "bad:Default "),
        xml(A).replace("</t:Types>", ""),
        xml(A)
            .replace("<t:Types ", "<t:Wrong ")
            .replace("</t:Types>", "</t:Wrong>"),
    ] {
        assert!(
            ContentTypeFragments::from_bytes(
                Arc::new(source.into_bytes()),
                SourceLimits::default()
            )
            .is_err()
        );
    }
}
