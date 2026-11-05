//! Entry transport preserves destination overlays and exact source spelling.
#![allow(clippy::unwrap_used, clippy::expect_used, reason = "test assertions")]
use litchi_opc::{
    OpcPackage, PackURI, PackageWriter, RelationshipFragmentEdit as Edit, RelationshipFragments,
    SourceLimits,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::sync::Arc;

const ROOT: &str = "<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships' xmlns:x='urn:extension'>";
const A: &str = "<r:Relationship TargetMode='External' Target='https://example.test/a' Type='urn:test' Id='a'/>";
const B: &str = "<r:Relationship Id=\"b\" Type='urn:test' Target='https://example.test/b' TargetMode='External'></r:Relationship>";
const C: &str = "<r:Relationship Type='urn:test' Id='c' TargetMode='External' Target='https://example.test/c'/>";
fn xml(body: &str) -> String {
    format!("\u{feff}{ROOT}{body}</r:Relationships><!--end-->")
}
fn owner() -> PackURI {
    PackURI::new("/custom/item.bin").unwrap()
}
fn catalog(xml: &str) -> RelationshipFragments {
    RelationshipFragments::from_bytes(
        owner(),
        Arc::new(xml.as_bytes().to_vec()),
        SourceLimits::default(),
    )
    .unwrap()
}
fn package(xml: &str) -> OpcPackage {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer.write_deflated("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#).unwrap();
    writer.write_deflated("custom/item.bin", b"opaque").unwrap();
    writer.write_deflated("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#).unwrap();
    writer
        .write_deflated("custom/_rels/item.bin.rels", xml.as_bytes())
        .unwrap();
    OpcPackage::from_vec(writer.finish_to_bytes().unwrap(), SourceLimits::default()).unwrap()
}
#[test]
fn removal_inverse_retains_exact_xml_and_unrelated_destination_overlays() {
    let original = xml(&format!("<?before keep?>{A}\n<!--gap-->{B}<x:extra/>{C}\n"));
    let mut package = package(&original);
    let limits = SourceLimits::default();
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let fragments = source.fragments(limits).unwrap();
    assert_eq!(fragments.iter().len(), 3);
    let removed = source
        .plan_fragment_edits(
            &[Edit {
                before: fragments.get("b"),
                after: None,
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    assert_eq!(removed.bytes(), original.replace(B, "").as_bytes());
    package
        .try_replace_relationships(&source, &removed, SourceLimits::default())
        .unwrap();
    let saved = PackageWriter::to_bytes(&package).unwrap();
    let mut package = OpcPackage::from_bytes(&saved, SourceLimits::default()).unwrap();
    let current = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let overlay = String::from_utf8(current.bytes().to_vec())
        .unwrap()
        .replace("before keep", "before overlay")
        .replace("https://example.test/a", "https://example.test/changed");
    package
        .try_replace_relationships_bytes(&current, Arc::new(overlay.into_bytes()), true, limits)
        .unwrap();
    let current = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let restored = current
        .plan_fragment_edits(
            &[Edit {
                before: None,
                after: fragments.get("b"),
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    let expected = original
        .replace("before keep", "before overlay")
        .replace("https://example.test/a", "https://example.test/changed");
    assert_eq!(restored.bytes(), expected.as_bytes());
    package
        .try_replace_relationships(&current, &restored, SourceLimits::default())
        .unwrap();
    let saved = PackageWriter::to_bytes(&package).unwrap();
    assert_eq!(
        ArchiveReader::new(&saved, soapberry_zip::office::ArchiveLimits::default())
            .unwrap()
            .read("custom/_rels/item.bin.rels")
            .unwrap(),
        expected.as_bytes()
    );
    assert_eq!(
        OpcPackage::from_bytes(&saved, SourceLimits::default())
            .unwrap()
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap()
            .bytes(),
        expected.as_bytes()
    );
}
#[test]
fn adjacent_suffix_inverse_precedes_unrelated_appended_entry() {
    let original = xml(&format!("<!--head-->{A}\n{B}<!--between-->{C}\n"));
    let fragments = catalog(&original);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let removed = source
        .plan_fragment_edits(
            &[
                Edit {
                    before: fragments.get("b"),
                    after: None,
                },
                Edit {
                    before: fragments.get("c"),
                    after: None,
                },
            ],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    let appended = A.replace("Id='a'", "Id='d'");
    let overlay = String::from_utf8(removed.bytes().to_vec())
        .unwrap()
        .replace(
            "</r:Relationships>",
            &format!("{appended}</r:Relationships>"),
        );
    let package = self::package(&overlay);
    let current = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    for reverse in [false, true] {
        let mut edits = [
            Edit {
                before: None,
                after: fragments.get("c"),
            },
            Edit {
                before: None,
                after: fragments.get("b"),
            },
        ];
        if reverse {
            edits.reverse();
        }
        let restored = current
            .plan_fragment_edits(&edits, limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        assert_eq!(
            restored.bytes(),
            original
                .replace(
                    "</r:Relationships>",
                    &format!("{appended}</r:Relationships>")
                )
                .as_bytes()
        );
    }
}
#[test]
fn replacement_and_identity_preserve_spelling() {
    let original = xml(&format!("{A}\n{B}\n{C}"));
    let changed = original.replace("https://example.test/b", "https://example.test/new");
    let before = catalog(&original);
    let after = catalog(&changed);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let noop = source
        .plan_fragment_edits(
            &[Edit {
                before: before.get("b"),
                after: before.get("b"),
            }],
            limits,
        )
        .unwrap();
    assert!(noop.is_noop());
    assert_eq!(
        noop.materialize(limits).unwrap().bytes(),
        original.as_bytes()
    );
    let plan = source
        .plan_fragment_edits(
            &[Edit {
                before: before.get("b"),
                after: after.get("b"),
            }],
            limits,
        )
        .unwrap();
    assert_eq!(plan.final_len(), changed.len());
    assert_eq!(plan.relationship_count(), 3);
    assert_eq!(
        plan.materialize(limits).unwrap().bytes(),
        changed.as_bytes()
    );
}
#[test]
fn first_entry_expands_aliased_empty_root() {
    let empty = format!("\u{feff}{} /><!--end-->", ROOT.trim_end_matches('>'));
    let after = catalog(&xml(A));
    let package = package(&empty);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let result = source
        .plan_fragment_edits(
            &[Edit {
                before: None,
                after: after.get("a"),
            }],
            limits,
        )
        .unwrap()
        .materialize(limits)
        .unwrap();
    assert_eq!(
        result.bytes(),
        xml(A)
            .replace("urn:extension'>", "urn:extension' >")
            .as_bytes()
    );
}
#[test]
fn stale_entry_context_and_duplicate_edits_leave_package_unchanged() {
    let original = xml(&format!("{A}{B}"));
    let before = catalog(&original);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    let changed =
        catalog(&original.replace("https://example.test/a", "https://example.test/stale"));
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: changed.get("a"),
                    after: None
                }],
                limits
            )
            .is_err()
    );
    let changed = catalog(&original.replace("urn:extension", "urn:other"));
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: before.get("a"),
                    after: changed.get("a")
                }],
                limits
            )
            .is_err()
    );
    let changed = catalog(&original.replace(
        "<r:Relationships ",
        "<r:Relationships xml:base='https://example.test/' ",
    ));
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: before.get("a"),
                    after: changed.get("a")
                }],
                limits
            )
            .is_err()
    );
    let edit = Edit {
        before: before.get("a"),
        after: None,
    };
    assert!(source.plan_fragment_edits(&[edit, edit], limits).is_err());
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: None,
                    after: before.get("a")
                }],
                limits
            )
            .is_err()
    );
    assert_eq!(
        package
            .source_relationships(&owner(), SourceLimits::default())
            .unwrap()
            .bytes(),
        original.as_bytes()
    );
}
#[test]
fn output_limits_are_admitted_before_materialization_and_rechecked() {
    let original = xml(A);
    let expanded = xml(&format!("{A}{B}"));
    let after = catalog(&expanded);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let edits = [Edit {
        before: None,
        after: after.get("b"),
    }];
    let limits = SourceLimits::default();
    let exact = SourceLimits::builder()
        .max_relationship_xml_bytes(expanded.len())
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
    let short = SourceLimits::builder()
        .max_relationship_xml_bytes(expanded.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(source.plan_fragment_edits(&edits, short).is_err());
    assert!(
        source
            .plan_fragment_edits(&edits, limits)
            .unwrap()
            .materialize(short)
            .is_err()
    );
    let count = SourceLimits::builder()
        .max_total_relationships(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(source.plan_fragment_edits(&edits, count).is_err());
}
#[test]
fn untrusted_fragments_reject_duplicate_ids_and_broken_xml() {
    for source in [
        xml(&format!("{A}{A}")),
        xml(A).replace("r:Relationship ", "bad:Relationship "),
        xml(A).replace("</r:Relationships>", ""),
    ] {
        assert!(
            RelationshipFragments::from_bytes(
                owner(),
                Arc::new(source.into_bytes()),
                SourceLimits::default()
            )
            .is_err()
        );
    }
}
#[test]
fn batched_replacements_keep_destination_slots_and_unrelated_gap_changes() {
    let original = xml(&format!("{A}\n{B}<!--gap-->{C}"));
    let changed = original.replace("https://example.test/", "https://example.test/new/");
    let before = catalog(&original);
    let after = catalog(&changed);
    let destination = original.replace("<!--gap-->", "<?retained overlay?>");
    let package = package(&destination);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let edits: Vec<_> = ["c", "a", "b"]
        .into_iter()
        .map(|id| Edit {
            before: before.get(id),
            after: after.get(id),
        })
        .collect();
    let result = source
        .plan_fragment_edits(&edits, SourceLimits::default())
        .unwrap()
        .materialize(SourceLimits::default())
        .unwrap();
    assert_eq!(
        result.bytes(),
        changed
            .replace("<!--gap-->", "<?retained overlay?>")
            .as_bytes()
    );
}
#[test]
fn every_deletion_subset_restores_exactly_in_either_batch_order() {
    let original = xml(&format!("<!--head-->{A}\n{B}<!--middle-->{C}\n"));
    let fragments = catalog(&original);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let limits = SourceLimits::default();
    for mask in 1..8 {
        let ids: Vec<_> = ["a", "b", "c"]
            .into_iter()
            .enumerate()
            .filter_map(|(index, id)| (mask & (1 << index) != 0).then_some(id))
            .collect();
        let edits: Vec<_> = ids
            .iter()
            .map(|id| Edit {
                before: fragments.get(id),
                after: None,
            })
            .collect();
        let removed = source
            .plan_fragment_edits(&edits, limits)
            .unwrap()
            .materialize(limits)
            .unwrap();
        let mut edits: Vec<_> = ids
            .iter()
            .map(|id| Edit {
                before: None,
                after: fragments.get(id),
            })
            .collect();
        for _ in 0..2 {
            edits.reverse();
            let restored = removed
                .plan_fragment_edits(&edits, limits)
                .unwrap()
                .materialize(limits)
                .unwrap();
            assert_eq!(restored.bytes(), original.as_bytes(), "mask {mask}");
        }
    }
}
#[test]
fn cycles_missing_anchors_and_wrong_owners_are_refused() {
    let empty = xml("");
    let package = package(&empty);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let ab = catalog(&xml(&format!("{A}{B}")));
    let ba = catalog(&xml(&format!("{B}{A}")));
    let limits = SourceLimits::default();
    assert!(
        source
            .plan_fragment_edits(
                &[
                    Edit {
                        before: None,
                        after: ab.get("a")
                    },
                    Edit {
                        before: None,
                        after: ba.get("b")
                    }
                ],
                limits
            )
            .is_err()
    );
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: None,
                    after: ab.get("a")
                }],
                limits
            )
            .is_err()
    );
    let foreign = RelationshipFragments::from_bytes(
        PackURI::new("/foreign.bin").unwrap(),
        Arc::new(xml(A).into_bytes()),
        limits,
    )
    .unwrap();
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: None,
                    after: foreign.get("a")
                }],
                limits
            )
            .is_err()
    );
    assert_eq!(source.bytes(), empty.as_bytes());
}
#[test]
fn event_and_aggregate_byte_bounds_apply_before_output() {
    let original = xml(A);
    let expanded = xml(&format!("{A}{B}{C}"));
    let fragments = catalog(&expanded);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let edits = [
        Edit {
            before: None,
            after: fragments.get("b"),
        },
        Edit {
            before: None,
            after: fragments.get("c"),
        },
    ];
    let plan = source
        .plan_fragment_edits(&edits, SourceLimits::default())
        .unwrap();
    let events = plan.event_upper_bound();
    let exact = SourceLimits::builder()
        .max_events(events)
        .unwrap()
        .max_total_relationship_xml_events(events)
        .unwrap()
        .max_total_relationship_xml_bytes(expanded.len())
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
    for limits in [
        SourceLimits::builder()
            .max_events(events - 1)
            .unwrap()
            .build()
            .unwrap(),
        SourceLimits::builder()
            .max_total_relationship_xml_events(events - 1)
            .unwrap()
            .build()
            .unwrap(),
        SourceLimits::builder()
            .max_total_relationship_xml_bytes(expanded.len() - 1)
            .unwrap()
            .build()
            .unwrap(),
    ] {
        assert!(source.plan_fragment_edits(&edits, limits).is_err());
    }
    let debug = format!("{fragments:?} {plan:?}");
    assert!(!debug.contains("example.test"));
    assert!(!debug.contains("Relationship Target"));
}
#[test]
fn stale_terminal_gap_cannot_relocate_restoration_after_an_unrelated_append() {
    let original = xml(&format!("{A}<!--gap-->{B}"));
    let fragments = catalog(&original);
    let appended = C;
    let destination = xml(&format!("{A}<!--changed-->{appended}"));
    let package = package(&destination);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    assert!(
        source
            .plan_fragment_edits(
                &[Edit {
                    before: None,
                    after: fragments.get("b")
                }],
                SourceLimits::default()
            )
            .is_err()
    );
    assert_eq!(source.bytes(), destination.as_bytes());
}
#[test]
fn no_op_rechecks_tightened_structural_limits() {
    let original = xml(A);
    let package = package(&original);
    let source = package
        .source_relationships(&owner(), SourceLimits::default())
        .unwrap();
    let plan = source
        .plan_fragment_edits(&[], SourceLimits::default())
        .unwrap();
    assert!(plan.is_noop());
    let shallow = SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(plan.materialize(shallow).is_err());
}
