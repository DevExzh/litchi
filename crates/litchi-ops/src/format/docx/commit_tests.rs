//! Transaction-result preservation at the unified/leaf boundary.

use std::io::{Cursor, Write as _};

use super::{Commit, Docx, WorkEdit};
use crate::format::docx::DocxOwner;
use crate::path::IntoPath;
use crate::props::{Key, PropValue};
use crate::transaction::SemanticLimits;
use crate::{File, IntoTarget, NewNode, OpenOptions, Placement, SourceMode, props};

const XML: &str = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>FIRST_PRIVATE</w:t></w:r><w:r><w:t>tail</w:t></w:r></w:p><!--between--><w:p><w:r><w:t>second</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#;

fn bytes() -> Vec<u8> {
    let mut writer = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, text) in [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#,
        ),
        ("word/document.xml", XML),
        ("opaque.xml", "<opaque>UNTOUCHED_PRIVATE</opaque>"),
    ] {
        writer.start_file(name, options).unwrap();
        writer.write_all(text.as_bytes()).unwrap();
    }
    writer.finish().unwrap().into_inner()
}

fn open(eager: bool) -> Docx {
    match crate::open_bytes_with(
        bytes(),
        &OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        }),
    )
    .unwrap()
    {
        File::Docx(handle) => handle,
        other => panic!("unexpected format: {other:?}"),
    }
}

fn xml(handle: &Docx) -> Vec<u8> {
    match handle.edit().unwrap() {
        crate::format::docx::DocxEditor::Edit(edit) => edit.source().unwrap().xml_bytes().to_vec(),
        crate::format::docx::DocxEditor::Source(edit) => {
            edit.source().unwrap().xml_bytes().to_vec()
        },
        crate::format::docx::DocxEditor::Unsigned(package) => {
            package.document_snapshot().unwrap().xml_bytes().to_vec()
        },
    }
}

fn apply(handle: &Docx, patch: &litchi_docx::document::Patch) -> Docx {
    match handle.owner() {
        DocxOwner::Eager(package) => {
            handle.adopt_eager_package(package.apply(patch).unwrap().into_package())
        },
        DocxOwner::Source(package) => {
            handle.adopt_source_package(package.apply(patch).unwrap().into_package())
        },
    }
}

fn check_commit(original: &Docx, commit: Commit) -> (Docx, litchi_docx::document::Patch) {
    assert!(!format!("{commit:?}").contains("PRIVATE"));
    let original_xml = xml(original);
    let (changed, patch, diagnostics) = commit.into_parts();
    assert!(diagnostics.changed());
    assert!(!patch.is_noop());
    assert_eq!(diagnostics.operations(), patch.operations().len());
    assert!(diagnostics.operations() > 0);
    assert_eq!(patch.source().xml_bytes(), original_xml);
    assert_eq!(patch.target().xml_bytes(), xml(&changed));
    assert_eq!(xml(&apply(original, &patch)), xml(&changed));
    assert_eq!(xml(&apply(&changed, &patch.inverse())), original_xml);
    assert_eq!(xml(original), original_xml);
    let replay = File::Docx(apply(original, &patch)).to_bytes().unwrap();
    assert_eq!(replay, File::Docx(changed.clone()).to_bytes().unwrap());
    (changed, patch)
}

#[test]
fn every_write_verb_keeps_its_leaf_patch_and_diagnostics() {
    for eager in [false, true] {
        let original = open(eager);
        check_commit(
            &original,
            original
                .set(
                    &"/body/p[1]/r[1]".into_target().unwrap(),
                    props! { "text" => "changed", "bold" => true },
                )
                .unwrap(),
        );
        let (commit, path) = original
            .add(
                "/body".into_path().unwrap(),
                NewNode::new("paragraph").prop("text", "third"),
            )
            .unwrap();
        assert_eq!(path.to_string(), "/body/p[3]");
        check_commit(&original, commit);
        let (commit, path) = original
            .add(
                "/body/p[2]".into_path().unwrap(),
                NewNode::new("run").prop("text", "new run"),
            )
            .unwrap();
        assert_eq!(path.to_string(), "/body/p[2]/r[2]");
        check_commit(&original, commit);
        check_commit(
            &original,
            original
                .remove("/body/p[1]/r[2]".into_target().unwrap())
                .unwrap(),
        );
        check_commit(
            &original,
            original
                .remove("/body/p[2]".into_target().unwrap())
                .unwrap(),
        );
        check_commit(
            &original,
            original
                .move_to("/body/p[1]".into_path().unwrap(), Placement::End)
                .unwrap(),
        );
        check_commit(
            &original,
            original
                .swap(
                    "/body/p[1]".into_path().unwrap(),
                    "/body/p[2]".into_path().unwrap(),
                )
                .unwrap(),
        );
        check_commit(
            &original,
            original
                .move_to(
                    "/body/p[1]/r[1]".into_path().unwrap(),
                    Placement::inside("/body/p[2]").unwrap(),
                )
                .unwrap(),
        );
        check_commit(
            &original,
            original
                .swap(
                    "/body/p[1]/r[1]".into_path().unwrap(),
                    "/body/p[2]/r[1]".into_path().unwrap(),
                )
                .unwrap(),
        );
    }
}

#[test]
fn exact_noops_still_have_real_patches_and_preserve_the_input_archive() {
    for eager in [false, true] {
        let original = open(eager);
        for commit in [
            original
                .set(&"/body/p[1]".into_target().unwrap(), props! {})
                .unwrap(),
            original
                .set(
                    &"/body/p[1]/r[1]".into_target().unwrap(),
                    props! { "text" => "FIRST_PRIVATE" },
                )
                .unwrap(),
            original
                .swap(
                    "/body/p[1]".into_path().unwrap(),
                    "/body/p[1]".into_path().unwrap(),
                )
                .unwrap(),
        ] {
            let (file, patch, diagnostics) = commit.into_parts();
            assert!(!diagnostics.changed());
            assert!(patch.is_noop());
            assert_eq!(diagnostics.operations(), patch.operations().len());
            assert_eq!(File::Docx(file).to_bytes().unwrap(), bytes());
            assert_eq!(
                File::Docx(apply(&original, &patch)).to_bytes().unwrap(),
                bytes()
            );
        }
    }
}

#[test]
fn ordered_commits_replay_and_reverse_against_their_own_intermediate_bases() {
    for eager in [false, true] {
        let original = open(eager);
        let (first, patch1) = check_commit(
            &original,
            original
                .set(
                    &"/body/p[1]/r[1]".into_target().unwrap(),
                    props! { "text" => "first step" },
                )
                .unwrap(),
        );
        let (second, patch2) = check_commit(
            &first,
            first.remove("/body/p[2]".into_target().unwrap()).unwrap(),
        );
        let replayed = apply(&apply(&original, &patch1), &patch2);
        assert_eq!(xml(&replayed), xml(&second));
        let restored = apply(&apply(&second, &patch2.inverse()), &patch1.inverse());
        assert_eq!(xml(&restored), xml(&original));
        let wrong_order = match original.owner() {
            DocxOwner::Eager(package) => package
                .apply(&patch2)
                .map(|_| ())
                .map_err(crate::Error::from),
            DocxOwner::Source(package) => package
                .apply(&patch2)
                .map(|_| ())
                .map_err(crate::Error::from),
        };
        assert!(
            wrong_order.is_err(),
            "a later patch cannot skip its earlier base"
        );
    }
}

#[test]
fn cancelled_leaf_operations_do_not_imply_a_changed_result() {
    for eager in [false, true] {
        let original = open(eager);
        let mut edit = WorkEdit::from(original.edit().unwrap()).unwrap();
        edit.replace_run_text(
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            "temporary".into(),
        )
        .unwrap();
        edit.replace_run_text(
            litchi_core::Position::new(0),
            litchi_core::Position::new(0),
            "FIRST_PRIVATE".into(),
        )
        .unwrap();
        let (file, patch, diagnostics) = edit.commit(&original).unwrap().into_parts();
        assert!(diagnostics.operations() > 0);
        assert!(!diagnostics.changed());
        assert!(patch.is_noop());
        assert_eq!(File::Docx(file).to_bytes().unwrap(), bytes());
    }
}

#[test]
fn failed_operation_does_not_damage_an_already_retained_commit() {
    for eager in [false, true] {
        let original = open(eager);
        let (first, patch) = check_commit(
            &original,
            original
                .set(
                    &"/body/p[1]/r[1]".into_target().unwrap(),
                    props! { "text" => "accepted" },
                )
                .unwrap(),
        );
        let before = xml(&first);
        assert!(
            first
                .set(
                    &"/body/p[1]/r[1]".into_target().unwrap(),
                    props! { "text" => "discard", "not.a.property" => true }
                )
                .is_err()
        );
        assert_eq!(xml(&first), before);
        assert_eq!(xml(&apply(&first, &patch.inverse())), xml(&original));
    }
}

#[test]
fn explicit_unsigning_is_outside_the_main_document_patch_scope() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/poi/test-data/xmldsign/hello-world-signed.docx"
    ));
    for eager in [false, true] {
        let options = OpenOptions::new().source_mode(if eager {
            SourceMode::Eager
        } else {
            SourceMode::Auto
        });
        let signed = crate::open_bytes_with(bytes.to_vec(), &options).unwrap();
        let File::Signed(signed) = signed else {
            panic!("signed fixture")
        };
        let unsigned = signed.unsign();
        let baseline = unsigned.to_bytes().unwrap();
        assert_ne!(baseline, bytes.as_slice());
        let File::Docx(handle) = unsigned else {
            panic!("unsigned DOCX")
        };
        let (file, patch, diagnostics) = handle
            .set(&"/body/p[1]".into_target().unwrap(), props! {})
            .unwrap()
            .into_parts();
        // These are leaf main-document facts, not proof that signature parts
        // remain unchanged relative to the original signed artifact.
        assert!(patch.is_noop());
        assert!(!diagnostics.changed());
        assert_eq!(diagnostics.operations(), 0);
        let output = File::Docx(file).to_bytes().unwrap();
        assert_eq!(output, baseline);
        assert!(!crate::open_bytes(output).unwrap().is_signed());
    }
}

#[cfg(any(unix, windows))]
#[test]
fn retained_source_commit_and_replay_still_fence_physical_source_changes() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    std::fs::write(&path, bytes()).unwrap();
    let File::Docx(original) = crate::open(&path).unwrap() else {
        panic!("DOCX source")
    };
    let (file, patch, _) = original
        .set(
            &"/body/p[1]/r[1]".into_target().unwrap(),
            props! { "text" => "changed" },
        )
        .unwrap()
        .into_parts();
    std::fs::write(&path, b"changed source revision").unwrap();
    assert_eq!(
        File::Docx(file).to_bytes().unwrap_err().code(),
        crate::Code::IoError
    );
    let DocxOwner::Source(package) = original.owner() else {
        panic!("source owner")
    };
    let error = crate::Error::from(package.apply(&patch).map(|_| ()).unwrap_err());
    assert_eq!(error.code(), crate::Code::IoError);
}

#[test]
fn empty_assignments_do_not_bypass_target_admission() {
    for eager in [false, true] {
        let original = open(eager);
        let error = original
            .set(&"/body".into_target().unwrap(), props! {})
            .unwrap_err();
        assert_eq!(error.code(), crate::Code::UnsupportedElement);
        let error = original
            .set(&"/body/p, r".into_target().unwrap(), props! {})
            .unwrap_err();
        assert_eq!(error.code(), crate::Code::UnsupportedCapability);
        assert_eq!(File::Docx(original).to_bytes().unwrap(), bytes());
    }
}

#[test]
fn paragraph_text_fanout_matches_scalar_bytes_and_retains_native_history() {
    for eager in [false, true] {
        for value in ["plain", "first\tline\nsecond", ""] {
            let original = open(eager);
            let commit = original
                .set(
                    &crate::query::Query::parse("/body/p")
                        .unwrap()
                        .into_target()
                        .unwrap(),
                    props! { "TEXT" => value },
                )
                .unwrap();
            let (changed, patch) = check_commit(&original, commit);
            assert_eq!(patch.operations().len(), 2);
            let first = original
                .set(
                    &"/body/p[1]".into_target().unwrap(),
                    props! { "text" => value },
                )
                .unwrap()
                .into_file();
            let scalar = first
                .set(
                    &"/body/p[2]".into_target().unwrap(),
                    props! { "text" => value },
                )
                .unwrap()
                .into_file();
            assert_eq!(
                File::Docx(changed.clone()).to_bytes().unwrap(),
                File::Docx(scalar).to_bytes().unwrap()
            );
            let (_, noop, diagnostics) = changed
                .set(
                    &crate::query::Query::parse("/body/p")
                        .unwrap()
                        .into_target()
                        .unwrap(),
                    props! { "text" => value },
                )
                .unwrap()
                .into_parts();
            assert!(noop.is_noop());
            assert!(!diagnostics.changed());
            assert_eq!(diagnostics.operations(), 0);
        }
    }
}

#[test]
fn paragraph_text_fanout_keeps_ordered_assignments_and_error_properties() {
    for eager in [false, true] {
        let original = open(eager);
        let before = xml(&original);
        let error = original
            .set(
                &crate::query::Query::parse("/body/p")
                    .unwrap()
                    .into_target()
                    .unwrap(),
                props! { "TEXT" => "bad\u{b}text" },
            )
            .unwrap_err();
        assert_eq!(error.property(), Some("TEXT"));
        assert_eq!(xml(&original), before);
        let commit = original
            .set(
                &crate::query::Query::parse("/body/p")
                    .unwrap()
                    .into_target()
                    .unwrap(),
                props! { "text" => "first", "TEXT" => "second", "align" => "center" },
            )
            .unwrap();
        let (changed, _) = check_commit(&original, commit);
        for position in [1, 2] {
            let file = File::Docx(changed.clone());
            assert_eq!(
                file.get(format!("/body/p[{position}]").as_str())
                    .unwrap()
                    .view()
                    .text()
                    .unwrap()
                    .as_deref(),
                Some("second")
            );
        }
    }
}

#[test]
fn immediate_assignment_borrows_and_deferred_assignment_owns_input() {
    let props = crate::Props::new().with(String::from("text"), "authored".repeat(1024));
    let (key, value) = props.iter().next().unwrap();
    let assignments = super::admit_assignments(super::TargetKind::Paragraph, &props).unwrap();
    assert_eq!(assignments[0].key.as_str().as_ptr(), key.as_str().as_ptr());
    assert_eq!(
        assignments[0].value.as_text().unwrap().as_ptr(),
        value.as_text().unwrap().as_ptr()
    );
    let owned = assignments.into_iter().next().unwrap().into_owned();
    drop(props);
    assert_eq!(owned.key.as_str(), "text");
    assert_eq!(owned.value.as_text().unwrap(), "authored".repeat(1024));
}

#[test]
fn fanout_write_proofs_load_at_most_one_body_per_admission() {
    use crate::format::docx::nav::BODY_LOADS;
    for eager in [true, false] {
        let handle = open(eager);
        // Paragraphs reuse the body index they were resolved from (record
        // 0251f); runs read it once per admission.
        for (selector, loads_expected) in [("/body/p", 0), ("/body/p/r", 1)] {
            let query = crate::query::Query::parse(selector).unwrap();
            let target = query.into_target().unwrap();
            let nodes = handle.resolve_targets(&target).unwrap();
            assert!(nodes.len() >= 2);
            let count = nodes.len();
            BODY_LOADS.with(|loads| loads.set(0));
            let targets = super::admit_targets(&nodes).unwrap();
            assert_eq!(targets.len(), count);
            BODY_LOADS.with(|loads| assert_eq!(loads.get(), loads_expected));
        }
    }
}

#[test]
fn find_main_story_admits_the_snapshot_budget_before_rewrite() {
    let original = open(true);
    let source_bytes = xml(&original).len();
    let replacement = "X".repeat(source_bytes);
    let props = crate::Props::from_props([
        (Key::from("find"), String::from("FIRST_PRIVATE")),
        (Key::from("replace"), replacement.clone()),
    ])
    .unwrap();
    let spec = crate::find::Spec::admit(&props).unwrap();
    let limits = SemanticLimits::builder()
        .max_snapshot_bytes(source_bytes.saturating_mul(2).saturating_sub(1))
        .unwrap()
        .build()
        .unwrap();
    let target = "/body".into_target().unwrap();
    let error = match original.find_replace(&target, &spec, &limits) {
        Ok(_) => panic!("an over-budget find rewrite was admitted"),
        Err(error) => error,
    };
    assert_eq!(error.code(), crate::Code::LimitExceeded);
}

/// The direct `/body/p[N]` text route of `File::set` (record 0305) publishes
/// exactly what the general route publishes, and fails exactly as it does:
/// the first paragraphs and the last of every DOCX fixture, one past the
/// last, values the text coercion accepts and refuses, and spellings and
/// owners it leaves to the general route.
#[test]
fn direct_paragraph_text_matches_the_general_route() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/docx");
    let mut fixtures: Vec<_> = std::fs::read_dir(&root)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| path.extension().is_some_and(|ext| ext == "docx"))
        .collect();
    fixtures.sort();
    let outcome = |result: crate::Result<Docx>| match result {
        Ok(handle) => Ok(File::Docx(handle).to_bytes().unwrap()),
        Err(error) => Err((error.code(), error.to_string())),
    };
    let mut direct = 0;
    for fixture in fixtures {
        let bytes = std::fs::read(&fixture).unwrap();
        for eager in [true, false] {
            if !eager && direct > 0 {
                // A source-backed owner always takes the general route.
                continue;
            }
            let options = OpenOptions::new().source_mode(if eager {
                SourceMode::Eager
            } else {
                SourceMode::Auto
            });
            let Ok(File::Docx(handle)) = crate::open_bytes_with(bytes.clone(), &options) else {
                continue;
            };
            let count = File::Docx(handle.clone())
                .query("/body > p")
                .map_or(0, Iterator::count);
            let values: [PropValue<'_>; 3] = ["new\ttext\nline".into(), 42.into(), PropValue::Null];
            let indices = (1..=count.min(6)).chain([count, count + 1]);
            for index in indices.filter(|index| *index > 0) {
                let mut paths = vec![format!("/body/p[{index}]")];
                if index == 1 {
                    paths.push(format!("/body/paragraph[{index}]"));
                    paths.push(format!("body/p[{index}]"));
                }
                for path in paths {
                    let target = path.as_str().into_target().unwrap();
                    for value in &values {
                        let props =
                            crate::Props::from_props([(Key::from("text"), value.clone())]).unwrap();
                        if handle.direct_paragraph(&target, &props).is_some() {
                            direct += 1;
                        }
                        let general = handle.set(&target, props.clone()).map(Commit::into_file);
                        let fast = handle.set_file(&target, props);
                        assert_eq!(
                            outcome(fast),
                            outcome(general),
                            "{} {path} {value:?} eager={eager}",
                            fixture.display()
                        );
                    }
                }
            }
        }
    }
    assert!(direct > 50, "the direct route ran {direct} times");
}
