//! Unified writes of OpenDocument presentations (record 0278): scene
//! splices through the leaf `Package::edit()`, exact no-ops, coalesced
//! batches, metadata, creation and atomic publication.

#![cfg(feature = "odp")]

use litchi_core::FileFormat;
use litchi_ops::capabilities::Verb;
use litchi_ops::{Code, File, NewNode, Placement, props};

fn fixture(relative: &str) -> std::path::PathBuf {
    std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

fn open(relative: &str) -> File {
    litchi_ops::open_bytes(std::fs::read(fixture(relative)).unwrap()).unwrap()
}

fn text(file: &File, path: &str) -> String {
    file.get(path)
        .unwrap_or_else(|error| panic!("{path}: {error}"))
        .view()
        .text()
        .unwrap()
        .map(|text| text.into_owned())
        .unwrap_or_default()
}

fn prop(file: &File, path: &str, key: &str) -> Option<String> {
    file.get(path)
        .unwrap()
        .view()
        .format()
        .unwrap()
        .find(|prop| prop.key() == key)
        .map(|prop| prop.to_string())
}

/// The `content.xml` of a published package.
fn content(file: &File) -> String {
    let package = litchi_odp::Package::from_bytes(&file.to_bytes().unwrap()).unwrap();
    package.presentation().unwrap().content_xml().to_owned()
}

fn reopen(file: &File) -> File {
    litchi_ops::open_bytes(file.to_bytes().unwrap()).unwrap()
}

#[test]
fn text_sets_rewrite_only_their_ranges() {
    let file = open("odf/corpus/impress-basic.odp");
    let before = file.to_bytes().unwrap();
    let source = content(&file);
    let updated = file
        .set("/slide[1]/shape[1]", props! { "text" => "New title" })
        .unwrap();
    assert_eq!(text(&updated, "/slide[1]/shape[1]"), "New title");
    // The input handle is unchanged.
    assert_eq!(file.to_bytes().unwrap(), before);
    assert_eq!(text(&file, "/slide[1]/shape[1]"), "OL");
    // Bytes before the title frame's text box are identical.
    let changed = content(&updated);
    let first_difference = source
        .bytes()
        .zip(changed.bytes())
        .position(|(left, right)| left != right)
        .unwrap();
    assert!(source[..first_difference].contains("presentation:class=\"title\""));
    let tail = 64;
    assert_eq!(
        &source[source.len() - tail..],
        &changed[changed.len() - tail..]
    );

    let reopened = reopen(&updated);
    assert_eq!(text(&reopened, "/slide[1]/shape[1]"), "New title");

    let paragraph = updated
        .set(
            "/slide[1]/shape[2]/paragraph[2]",
            props! { "text" => "Two  spaced" },
        )
        .unwrap();
    assert_eq!(
        text(&paragraph, "/slide[1]/shape[2]/paragraph[2]"),
        "Two  spaced"
    );
    assert_eq!(text(&paragraph, "/slide[1]/shape[2]/paragraph[1]"), "Eins");
    let run = paragraph
        .set(
            "/slide[1]/shape[2]/paragraph[3]/run[1]",
            props! { "text" => "Drei" },
        )
        .unwrap();
    assert_eq!(text(&run, "/slide[1]/shape[2]/paragraph[3]"), "Drei");
    let notes = run
        .set("/slide[1]/notes", props! { "text" => "Speaker\nnotes" })
        .unwrap();
    assert_eq!(text(&notes, "/slide[1]/notes"), "Speaker\nnotes");
    assert_eq!(
        text(&reopen(&notes), "/slide[1]/notes/paragraph[2]"),
        "notes"
    );
}

#[test]
fn exact_noops_publish_the_source_bytes() {
    let file = open("odf/corpus/impress-basic.odp");
    let before = file.to_bytes().unwrap();
    let same = file
        .set("/slide[1]/shape[1]", props! { "text" => "OL" })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), before);
    let x = prop(&file, "/slide[1]/shape[1]", "x").unwrap();
    let same = file
        .set("/slide[1]/shape[1]", props! { "x" => x.as_str() })
        .unwrap();
    assert_eq!(same.to_bytes().unwrap(), before);
    let commit = file
        .edit(|tx| {
            tx.set(
                "/slide[1]/shape[2]/paragraph[1]",
                props! { "text" => "Eins" },
            )?;
            Ok(())
        })
        .unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.file().to_bytes().unwrap(), before);
}

#[test]
fn geometry_names_and_cells_set() {
    let file = open("odf/odp/cellspan.odp");
    let updated = file
        .set(
            "/slide[1]/table[1]",
            props! { "name" => "Grid", "width" => "12cm", "x" => "1in" },
        )
        .unwrap();
    assert_eq!(
        prop(&updated, "/slide[1]/table[1]", "name").as_deref(),
        Some("Grid")
    );
    assert_eq!(
        prop(&updated, "/slide[1]/table[1]", "width").as_deref(),
        Some("12cm")
    );
    assert_eq!(
        prop(&updated, "/slide[1]/table[1]", "x").as_deref(),
        Some("2.54cm")
    );
    let cell = updated
        .set(
            "/slide[1]/table[1]/tr[1]/tc[1]",
            props! { "text" => "Cell one" },
        )
        .unwrap();
    assert_eq!(
        text(&reopen(&cell), "/slide[1]/table[1]/tr[1]/tc[1]"),
        "Cell one"
    );
    // Field runs and unknown keys are typed refusals.
    let error = file
        .set("/slide[1]/table[1]", props! { "fill" => "FF0000" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
    let error = file
        .set("/slide[1]/table[1]", props! { "width" => "-1cm" })
        .unwrap_err();
    assert_eq!(error.code(), Code::InvalidValue);
}

#[test]
fn slides_shapes_and_paragraphs_add_remove_move_and_swap() {
    let file = open("odf/odp/tdf169979.odp");
    assert_eq!(file.query("slide").unwrap().count(), 3);
    let added = file
        .add("/", NewNode::new("slide").prop("name", "Appendix"))
        .unwrap();
    assert_eq!(added.path.to_string(), "/slide[4]");
    let file4 = added.file;
    assert_eq!(
        prop(&file4, "/slide[4]", "name").as_deref(),
        Some("Appendix")
    );
    let shape = file4
        .add(
            "/slide[4]",
            NewNode::new("shape")
                .prop("text", "Hello\nWorld")
                .prop("width", "10cm")
                .prop("height", "3cm")
                .prop("name", "Box"),
        )
        .unwrap();
    assert_eq!(shape.path.to_string(), "/slide[4]/shape[1]");
    assert_eq!(text(&shape.file, "/slide[4]/shape[1]"), "Hello\nWorld");
    let paragraph = shape
        .file
        .add(
            "/slide[4]/shape[1]",
            NewNode::new("paragraph")
                .prop("text", "Middle")
                .at(Placement::Position(litchi_core::Position::new(1))),
        )
        .unwrap();
    assert_eq!(
        paragraph.path.to_string(),
        "/slide[4]/shape[1]/paragraph[2]"
    );
    assert_eq!(
        text(&paragraph.file, "/slide[4]/shape[1]"),
        "Hello\nMiddle\nWorld"
    );
    let removed = paragraph
        .file
        .remove("/slide[4]/shape[1]/paragraph[1]")
        .unwrap();
    assert_eq!(text(&removed, "/slide[4]/shape[1]"), "Middle\nWorld");

    let moved = removed.move_to("/slide[4]", Placement::Start).unwrap();
    assert_eq!(
        prop(&moved, "/slide[1]", "name").as_deref(),
        Some("Appendix")
    );
    let swapped = moved.swap("/slide[1]", "/slide[2]").unwrap();
    assert_eq!(
        prop(&swapped, "/slide[2]", "name").as_deref(),
        Some("Appendix")
    );
    let gone = swapped.remove("/slide[2]").unwrap();
    assert_eq!(gone.query("slide").unwrap().count(), 3);
    let reopened = reopen(&gone);
    assert_eq!(reopened.query("slide").unwrap().count(), 3);
    assert_eq!(
        text(&reopened, "/slide[1]/shape[1]"),
        text(&file, "/slide[1]/shape[1]")
    );

    // A shape removal and z-order move on one slide.
    let shape_removed = gone.remove("/slide[1]/shape[1]").unwrap();
    assert!(shape_removed.find("/slide[1]/shape[1]").unwrap().is_none());

    // The last slide stays.
    let single = open("odf/corpus/impress-basic.odp");
    assert_eq!(
        single.remove("/slide[1]").unwrap_err().code(),
        Code::InvalidArgument
    );
}

#[test]
fn metadata_reads_and_sets() {
    let file = open("odf/corpus/impress-basic.odp");
    let updated = file
        .set(
            "/metadata",
            props! { "title" => "Deck", "subject" => "Unified", "keywords" => litchi_ops::PropValue::Null },
        )
        .unwrap();
    let metadata = reopen(&updated).metadata().unwrap().unwrap();
    assert_eq!(metadata.title.as_deref(), Some("Deck"));
    assert_eq!(metadata.subject.as_deref(), Some("Unified"));
    assert_eq!(metadata.author.as_deref(), Some("pair3"));
    let error = file
        .set("/metadata", props! { "company" => "Litchi" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedCapability);
    let error = file
        .set("/metadata", props! { "nonsense" => "x" })
        .unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedProperty);
}

#[test]
fn batches_coalesce_into_one_leaf_commit() {
    let file = open("odf/corpus/impress-basic.odp");
    let commit = file
        .edit(|tx| {
            for (index, text) in ["a", "b", "c", "d", "e"].iter().enumerate() {
                tx.set(
                    format!("/slide[1]/shape[2]/paragraph[{}]", index + 1).as_str(),
                    props! { "text" => *text },
                )?;
            }
            tx.set(
                "/slide[1]/shape[1]",
                props! { "name" => "Head", "width" => "20cm" },
            )?;
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().operations(), 6);
    assert_eq!(commit.diagnostics().commits(), 1);
    assert_eq!(text(commit.file(), "/slide[1]/shape[2]/paragraph[5]"), "e");
    let restored = commit.patch().inverse().apply(commit.file()).unwrap();
    assert_eq!(restored.to_bytes().unwrap(), file.to_bytes().unwrap());

    // A story set and then one of its paragraphs: the second resolves
    // against the first's result.
    let commit = file
        .edit(|tx| {
            tx.set("/slide[1]/shape[2]", props! { "text" => "x\ny" })?;
            tx.set("/slide[1]/shape[2]/paragraph[2]", props! { "text" => "z" })?;
            Ok(())
        })
        .unwrap();
    assert_eq!(commit.diagnostics().commits(), 2);
    assert_eq!(text(commit.file(), "/slide[1]/shape[2]"), "x\nz");

    // A failed batch leaves the input unchanged and names the item.
    let before = file.to_bytes().unwrap();
    let error = file
        .edit(|tx| {
            tx.set("/slide[1]/shape[1]", props! { "text" => "ok" })?;
            tx.set("/slide[1]/shape[9]", props! { "text" => "missing" })?;
            Ok(())
        })
        .unwrap_err();
    assert_eq!(error.item(), Some(1));
    assert_eq!(file.to_bytes().unwrap(), before);
}

#[test]
fn create_save_and_flat_refusals() {
    let created = litchi_ops::create(FileFormat::Odp).unwrap();
    assert_eq!(created.query("slide").unwrap().count(), 1);
    let added = created
        .add(
            "/slide[1]",
            NewNode::new("shape")
                .prop("text", "Created")
                .prop("width", "8cm")
                .prop("height", "2cm"),
        )
        .unwrap();
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("created.odp");
    added.file.save_as(&path).unwrap();
    let reopened = litchi_ops::open(&path).unwrap();
    assert_eq!(text(&reopened, "/slide[1]/shape[1]"), "Created");
    // A path-opened file saves to its origin atomically.
    let edited = reopened
        .set("/slide[1]/shape[1]", props! { "text" => "Saved" })
        .unwrap();
    edited.save().unwrap();
    assert_eq!(
        text(&litchi_ops::open(&path).unwrap(), "/slide[1]/shape[1]"),
        "Saved"
    );
    assert_eq!(
        std::fs::read_dir(directory.path()).unwrap().count(),
        1,
        "no staging file is left"
    );
    // An existing destination is kept when the caller refuses overwrites.
    let saved = std::fs::read(&path).unwrap();
    let refused = created.save_with(
        litchi_ops::SaveTo::Path(&path),
        &litchi_ops::SaveOptions::new().overwrite(litchi_ops::Overwrite::Refuse),
    );
    assert!(refused.is_err());
    assert_eq!(std::fs::read(&path).unwrap(), saved);

    for (index, relative) in [
        "libreoffice-core/sd/qa/unit/data/draw-image-link.fodp",
        "libreoffice-core/sd/qa/unit/data/draw-object-link.fodp",
    ]
    .into_iter()
    .enumerate()
    {
        let source_path = fixture(relative);
        let flat = litchi_ops::open(&source_path).unwrap();
        // Record 0292: a flat presentation edits its packaged projection and
        // republishes the edit into its flat bytes (`flat_odf_unified.rs`).
        // The projection admission now handles these producers' empty
        // master-page names; prove that the page rename is real and
        // reversible, rather than accepting a false no-op.
        let capabilities = flat.capabilities();
        assert!(capabilities.verb(Verb::Set).is_supported());
        assert!(capabilities.verb(Verb::Get).is_supported());
        let before = flat.to_bytes().unwrap();
        let original = prop(&flat, "/slide[1]", "name").unwrap();
        let renamed = flat.set("/slide[1]", props! { "name" => "x" }).unwrap();
        assert_eq!(prop(&renamed, "/slide[1]", "name").as_deref(), Some("x"));
        assert_ne!(renamed.to_bytes().unwrap(), before);
        let restored = renamed
            .set("/slide[1]", props! { "name" => original.as_str() })
            .unwrap();
        assert_eq!(prop(&restored, "/slide[1]", "name"), Some(original));
        assert_eq!(restored.to_bytes().unwrap(), before);
        let copy = directory.path().join(format!("copy-{index}.fodp"));
        flat.save_as(&copy).unwrap();
        assert_eq!(
            std::fs::read(&copy).unwrap(),
            std::fs::read(&source_path).unwrap()
        );
    }
}
