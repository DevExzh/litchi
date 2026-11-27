//! Equation reads and edits, on an Equation Editor 3.0 object PowerPoint
//! saved (Apache POI `slideshow/42520.ppt`) and on presentations this
//! crate's writer authors.

use std::collections::BTreeMap;
use std::io::Cursor;

use litchi_core::Position;
use litchi_core::edit::RefusalKind;
use litchi_equation::Limits;

use super::read::{self, EQUATION_NATIVE};
use super::{Equation, Error, Source};
use crate::Package;

const FIXTURE: &[u8] = include_bytes!("../../../../test-data/poi/test-data/slideshow/42520.ppt");
/// The fixture's equation: slide 16, frame 17424, object 61, persist 21.
const SLIDE: usize = 15;
const SHAPE: u32 = 17424;

fn open(bytes: &[u8]) -> Package {
    Package::from_bytes(bytes).unwrap()
}

fn equations(package: &Package) -> Vec<Equation> {
    package.presentation().unwrap().equations().unwrap()
}

fn latex(equation: &Equation) -> String {
    equation.latex(&Limits::DEFAULT).unwrap()
}

/// Every live persisted record, by persist identifier.
fn records(package: &Package) -> BTreeMap<u32, Vec<u8>> {
    let presentation = package.presentation().unwrap();
    presentation
        .inner
        .persist_mapping
        .get_persist_ids()
        .into_iter()
        .map(|id| {
            (
                id,
                presentation.persisted_record_bytes(id).unwrap().to_vec(),
            )
        })
        .collect()
}

/// Every stream of the package's compound file.
fn streams(package: &Package) -> BTreeMap<String, Vec<u8>> {
    let bytes = package.to_bytes().unwrap();
    let mut file = litchi_cfb::OleFile::open(Cursor::new(bytes), Default::default()).unwrap();
    file.list_streams()
        .into_iter()
        .map(|path| {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            (path.join("/"), file.open_stream(&refs).unwrap())
        })
        .collect()
}

fn slide_persist(package: &Package, slide: usize) -> u32 {
    package.presentation().unwrap().slide_directory().entries()[slide].persist_id()
}

fn document_persist(package: &Package) -> u32 {
    package
        .presentation()
        .unwrap()
        .slide_directory()
        .document_persist_id()
}

/// The stored compound file of an equation, decompressed.
fn storage_streams(package: &Package, persist_id: u32) -> BTreeMap<String, Vec<u8>> {
    let storage = package
        .presentation()
        .unwrap()
        .ole_storage(persist_id)
        .unwrap()
        .unwrap();
    let bytes = storage.decompressed_bytes(read::MAX_STORAGE_BYTES).unwrap();
    let mut file = litchi_cfb::OleFile::open(Cursor::new(bytes), Default::default()).unwrap();
    file.list_streams()
        .into_iter()
        .map(|path| {
            let refs = path.iter().map(String::as_str).collect::<Vec<_>>();
            (path.join("/"), file.open_stream(&refs).unwrap())
        })
        .collect()
}

/// The picture-store entries of the live document: (type, size, refs,
/// offset).
fn store_entries(package: &Package) -> Vec<(u8, u32, u32, u32)> {
    let presentation = package.presentation().unwrap();
    let id = presentation.slide_directory().document_persist_id();
    let mut document =
        super::tree::parse(presentation.persisted_record_bytes(id).unwrap()).unwrap();
    let group = super::object::drawing_group(&mut document).unwrap();
    group.expand_all(0).unwrap();
    let Some(store) = group.child(0xF001) else {
        return Vec::new();
    };
    store
        .children()
        .iter()
        .map(|entry| {
            let payload = entry.payload().unwrap();
            let field = |at: usize| u32::from_le_bytes(payload[at..at + 4].try_into().unwrap());
            (payload[0], field(20), field(24), field(28))
        })
        .collect()
}

/// A presentation with `slides` slides, each with one text box, authored
/// by this crate's writer.
fn authored(slides: usize) -> Vec<u8> {
    let mut writer = crate::writer::Builder::new();
    for index in 0..slides {
        let slide = writer.add_slide().unwrap();
        writer
            .add_textbox(slide, 10, 10, 240, 40, &format!("slide {index}"))
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    crate::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

fn mtef_version(equation: &Equation) -> u8 {
    litchi_equation::read_mtef(&equation.equation_native().unwrap(), &Limits::DEFAULT)
        .unwrap()
        .info()
        .unwrap()
        .version
}

#[test]
fn reads_the_equation_editor_object_powerpoint_saved() {
    let package = open(FIXTURE);
    let found = equations(&package);
    assert_eq!(found.len(), 1);
    let equation = &found[0];
    assert_eq!(equation.slide(), SLIDE);
    assert_eq!(equation.shape_id(), SHAPE);
    assert_eq!(equation.object_id(), 61);
    assert_eq!(equation.persist_id(), 21);
    assert_eq!(equation.program_id(), Some("Equation.3"));
    assert!(equation.is_compressed());
    assert!(!equation.is_grouped());
    assert_eq!(
        equation.bounds(),
        Some(super::Bounds::from_edges(2009, 2208, 3831, 2434))
    );
    assert_eq!(equation.equation_native().unwrap().len(), 113);
    assert_eq!(&equation.equation_native().unwrap()[..2], &[28, 0]);
    assert_eq!(mtef_version(equation), 3);
    assert_eq!(latex(equation), "y= 0.317+ (- .008•DSO)");
}

#[test]
fn an_unchanged_edit_publishes_the_source() {
    let package = open(FIXTURE);
    let edit = package.edit_equations().unwrap();
    assert_eq!(edit.equations().unwrap(), equations(&package));
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), FIXTURE);
}

#[test]
fn replacing_keeps_the_object_kind_compression_and_other_records() {
    let package = open(FIXTURE);
    let before_records = records(&package);
    let before_streams = storage_streams(&package, 21);
    let before_store = store_entries(&package);
    let mut edit = package.edit_equations().unwrap();
    let written = edit
        .set(
            Position::new(SLIDE),
            SHAPE,
            Source::Latex(r"\frac{a+b}{2}"),
            &Limits::DEFAULT,
        )
        .unwrap();
    assert_eq!(written.shape_id(), SHAPE);
    assert!(written.losses().is_empty());
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let after = commit.package();

    let found = equations(after);
    assert_eq!(found.len(), 1);
    let equation = &found[0];
    assert_eq!(equation.shape_id(), SHAPE);
    assert_eq!(equation.program_id(), Some("Equation.3"));
    assert_eq!(mtef_version(equation), 3);
    assert!(equation.is_compressed());
    assert_eq!(latex(equation), r"\frac{a+ b}{2}");
    // The frame keeps its corner and grows with the taller equation.
    let bounds = equation.bounds().unwrap();
    assert_eq!((bounds.left, bounds.top), (2009, 2208));
    assert!(bounds.height() > 2434 - 2208);

    // The storage keeps every stream but the equation and the stale cached
    // presentation.
    let after_streams = storage_streams(after, 21);
    assert!(!after_streams.contains_key("\u{2}OlePres000"));
    for (name, bytes) in &before_streams {
        if name != EQUATION_NATIVE && name != "\u{2}OlePres000" {
            assert_eq!(after_streams.get(name), Some(bytes), "{name:?}");
        }
    }

    // Only the slide, the document (picture store) and the storage change.
    let after_records = records(after);
    let changed = [
        slide_persist(&package, SLIDE),
        document_persist(&package),
        21,
    ];
    assert_eq!(
        before_records.keys().collect::<Vec<_>>(),
        after_records.keys().collect::<Vec<_>>()
    );
    for (id, bytes) in &before_records {
        if !changed.contains(id) {
            assert_eq!(after_records.get(id), Some(bytes), "persist {id}");
        }
    }

    // The old WMF preview (entry 7, one reference) is freed and its BLIP
    // removed; the next BLIP moves down; the PNG preview is appended.
    let after_store = store_entries(after);
    assert_eq!(after_store.len(), before_store.len() + 1);
    assert_eq!(after_store[6], (0, 0, 0, 0));
    assert_eq!(after_store[7].3, before_store[6].3);
    assert_eq!(after_store.last().unwrap().0, 6);
    assert_eq!(after_store.last().unwrap().2, 1);
    for index in 0..6 {
        assert_eq!(after_store[index], before_store[index]);
    }
}

#[test]
fn replacing_with_raw_mtef_of_another_version_is_refused() {
    let package = open(FIXTURE);
    let mathtype = litchi_equation::latex_to_equation_native(
        "x",
        &litchi_equation::MtefTarget::MathType,
        &Limits::DEFAULT,
    )
    .unwrap();
    let mut edit = package.edit_equations().unwrap();
    let error = edit
        .set(
            Position::new(SLIDE),
            SHAPE,
            Source::Mtef(mathtype.bytes()),
            &Limits::DEFAULT,
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::InvalidValue),
        "{error:?}"
    );
    // The failed operation left nothing staged.
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn replacing_with_raw_mtef_of_the_same_version_stores_it_exactly() {
    let package = open(FIXTURE);
    let native = litchi_equation::latex_to_equation_native(
        "z^2",
        &litchi_equation::MtefTarget::EquationEditor3,
        &Limits::DEFAULT,
    )
    .unwrap();
    let mut edit = package.edit_equations().unwrap();
    edit.set(
        Position::new(SLIDE),
        SHAPE,
        Source::Mtef(native.bytes()),
        &Limits::DEFAULT,
    )
    .unwrap();
    let after = edit.commit().unwrap().into_parts().0;
    assert_eq!(
        &*equations(&after)[0].equation_native().unwrap(),
        native.bytes()
    );
}

#[test]
fn removing_deletes_the_shape_object_storage_and_preview() {
    let package = open(FIXTURE);
    let before_records = records(&package);
    let before_store = store_entries(&package);
    let before_pictures = streams(&package)["Pictures"].clone();
    let mut edit = package.edit_equations().unwrap();
    edit.remove(Position::new(SLIDE), SHAPE).unwrap();
    let after = edit.commit().unwrap().into_parts().0;

    assert!(equations(&after).is_empty());
    let presentation = after.presentation().unwrap();
    let objects = presentation.ole_objects().unwrap().unwrap();
    assert!(objects.get(61).is_none());
    assert!(objects.get(57).is_some());
    let slide = presentation
        .persisted_record_bytes(slide_persist(&after, SLIDE))
        .unwrap();
    assert!(
        read::object_shapes(slide)
            .unwrap()
            .iter()
            .all(|shape| shape.shape_id != SHAPE)
    );

    // No object references the storage any more. (An incremental save
    // cannot drop a persist identifier from the directories of earlier
    // edits, so its old record stays reachable by identifier.)
    assert!(
        objects
            .objects
            .iter()
            .all(|object| object.persist_id() != 21)
    );
    let after_records = records(&after);
    let changed = [
        slide_persist(&package, SLIDE),
        document_persist(&package),
        21,
    ];
    for (id, bytes) in &before_records {
        if !changed.contains(id) {
            assert_eq!(after_records.get(id), Some(bytes), "persist {id}");
        }
    }

    // The preview's entry is emptied and its 401-byte BLIP removed.
    let after_store = store_entries(&after);
    assert_eq!(after_store.len(), before_store.len());
    assert_eq!(after_store[6], (0, 0, 0, 0));
    assert_eq!(after_store[7].3, before_store[6].3);
    let after_pictures = &streams(&after)["Pictures"];
    assert_eq!(after_pictures.len(), before_pictures.len() - 401);
    let (start, end) = (before_store[6].3 as usize, before_store[6].3 as usize + 401);
    assert_eq!(&after_pictures[..start], &before_pictures[..start]);
    assert_eq!(&after_pictures[start..], &before_pictures[end..]);
}

#[test]
fn adding_writes_a_mathtype_object_with_a_png_preview() {
    let package = open(FIXTURE);
    let before_records = records(&package);
    let before_store = store_entries(&package);
    let mut edit = package.edit_equations().unwrap();
    let written = edit
        .add(
            Position::new(0),
            576,
            1152,
            24.0,
            Source::Latex(r"E=mc^{2}"),
            &Limits::DEFAULT,
        )
        .unwrap();
    let after = edit.commit().unwrap().into_parts().0;

    let found = equations(&after);
    assert_eq!(found.len(), 2);
    let added = found.iter().find(|equation| equation.slide() == 0).unwrap();
    assert_eq!(added.shape_id(), written.shape_id());
    assert_eq!(added.program_id(), Some("Equation.DSMT4"));
    assert_eq!(mtef_version(added), 5);
    assert!(added.is_compressed());
    assert_eq!(latex(added), "E= m{c}^{2}");
    let bounds = added.bounds().unwrap();
    assert_eq!((bounds.left, bounds.top), (576, 1152));
    assert!(bounds.width() > bounds.height() && bounds.height() > 100);
    // The untouched equation keeps its bytes.
    let original = found
        .iter()
        .find(|equation| equation.slide() == SLIDE)
        .unwrap();
    assert_eq!(original, &equations(&package)[0]);

    // The storage holds the MathType registration streams.
    let stored = storage_streams(&after, added.persist_id());
    assert!(
        stored["\u{1}CompObj"]
            .windows(14)
            .any(|window| window == b"Equation.DSMT4")
    );
    assert_eq!(stored[EQUATION_NATIVE], &*added.equation_native().unwrap());
    // A new persist identifier; every source record keeps its bytes but the
    // slide and the document.
    let after_records = records(&after);
    assert!(!before_records.contains_key(&added.persist_id()));
    let changed = [slide_persist(&package, 0), document_persist(&package)];
    for (id, bytes) in &before_records {
        if !changed.contains(id) {
            assert_eq!(after_records.get(id), Some(bytes), "persist {id}");
        }
    }
    // The preview is a new PNG entry.
    let after_store = store_entries(&after);
    assert_eq!(&after_store[..before_store.len()], &before_store[..]);
    assert_eq!(after_store.len(), before_store.len() + 1);
    let picture = streams(&after)["Pictures"].clone();
    let entry = after_store.last().unwrap();
    assert_eq!(entry.0, 6);
    let blip = &picture[entry.3 as usize..(entry.3 + entry.1) as usize];
    assert_eq!(&blip[25..33], b"\x89PNG\r\n\x1a\n");
}

#[test]
fn authored_presentations_take_add_set_and_remove() {
    let source = authored(2);
    let package = open(&source);
    assert!(equations(&package).is_empty());
    let mut edit = package.edit_equations().unwrap();
    let first = edit
        .add(
            Position::new(1),
            100,
            100,
            18.0,
            Source::Latex("x+1"),
            &Limits::DEFAULT,
        )
        .unwrap();
    let second = edit
        .add(
            Position::new(1),
            100,
            600,
            18.0,
            Source::Omml("<m:oMath><m:r><m:t>y</m:t></m:r></m:oMath>"),
            &Limits::DEFAULT,
        )
        .unwrap();
    assert_ne!(first.shape_id(), second.shape_id());
    assert_eq!(edit.equations().unwrap().len(), 2);
    let after = edit.commit().unwrap().into_parts().0;
    let found = equations(&after);
    assert_eq!(found.iter().map(latex).collect::<Vec<_>>(), ["x+ 1", "y"]);

    // Replacing one leaves the other's bytes alone.
    let mut edit = after.edit_equations().unwrap();
    edit.set(
        Position::new(1),
        first.shape_id(),
        Source::Latex(r"\sqrt{x}"),
        &Limits::DEFAULT,
    )
    .unwrap();
    let updated = edit.commit().unwrap().into_parts().0;
    let found_updated = equations(&updated);
    assert_eq!(latex(&found_updated[0]), r"\sqrt{x}");
    assert_eq!(found_updated[1], found[1]);

    // Removing both leaves no equation and no object.
    let mut edit = updated.edit_equations().unwrap();
    edit.remove(Position::new(1), first.shape_id()).unwrap();
    edit.remove(Position::new(1), second.shape_id()).unwrap();
    let removed = edit.commit().unwrap().into_parts().0;
    assert!(equations(&removed).is_empty());
    assert!(
        removed
            .presentation()
            .unwrap()
            .ole_objects()
            .unwrap()
            .is_none_or(|objects| objects.objects.is_empty())
    );
    assert!(
        store_entries(&removed)
            .iter()
            .all(|entry| *entry == (0, 0, 0, 0))
    );
    assert!(streams(&removed).get("Pictures").is_none_or(Vec::is_empty));
}

#[test]
fn an_edit_patch_replays_and_inverts() {
    let package = open(FIXTURE);
    let mut edit = package.edit_equations().unwrap();
    edit.set(
        Position::new(SLIDE),
        SHAPE,
        Source::Latex("q"),
        &Limits::DEFAULT,
    )
    .unwrap();
    let (after, patch) = edit.commit().unwrap().into_parts();
    let replayed = package.apply_equations(&patch).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        after.to_bytes().unwrap()
    );
    let restored = after.apply_equations(&patch.inverse()).unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), FIXTURE);
    assert!(matches!(
        after.apply_equations(&patch),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn missing_targets_and_bad_sources_are_refused_without_staging() {
    let package = open(FIXTURE);
    let mut edit = package.edit_equations().unwrap();
    let not_found = |error: Error| matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::NotFound);
    assert!(not_found(edit.remove(Position::new(SLIDE), 1).unwrap_err()));
    assert!(not_found(
        edit.remove(Position::new(99), SHAPE).unwrap_err()
    ));
    assert!(not_found(
        edit.set(
            Position::new(0),
            SHAPE,
            Source::Latex("x"),
            &Limits::DEFAULT
        )
        .unwrap_err()
    ));
    assert!(not_found(
        edit.add(
            Position::new(99),
            0,
            0,
            12.0,
            Source::Latex("x"),
            &Limits::DEFAULT
        )
        .unwrap_err()
    ));
    // A text shape is not an equation.
    assert!(not_found(
        edit.remove(Position::new(SLIDE), 17422).unwrap_err()
    ));
    assert!(matches!(
        edit.set(
            Position::new(SLIDE),
            SHAPE,
            Source::Latex(r"\frac{"),
            &Limits::DEFAULT
        )
        .unwrap_err(),
        Error::Equation(litchi_equation::Error::Malformed { .. })
    ));
    assert!(matches!(
        edit.add(
            Position::new(0),
            0,
            0,
            12.0,
            Source::Mtef(b"not mtef"),
            &Limits::DEFAULT
        )
        .unwrap_err(),
        Error::Equation(_)
    ));
    let tight = Limits::new(64, 1_000, 8).unwrap();
    let deep = format!("{}x{}", r"\frac{1}{".repeat(40), "}".repeat(40));
    assert!(matches!(
        edit.add(Position::new(0), 0, 0, 12.0, Source::Latex(&deep), &tight)
            .unwrap_err(),
        Error::Equation(_)
    ));
    // A frame beyond the slide coordinates.
    assert!(matches!(
        edit.add(
            Position::new(0),
            32_760,
            0,
            12.0,
            Source::Latex("x"),
            &Limits::DEFAULT
        )
        .unwrap_err(),
        Error::Refused(_)
    ));
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn every_slide_takes_a_new_equation_with_an_unused_shape_identifier() {
    // Some slides of this deck use shape identifiers of a cluster the
    // drawing group assigns to another drawing; those get a new cluster.
    let package = open(FIXTURE);
    let count = package.presentation().unwrap().slide_count();
    let mut edit = package.edit_equations().unwrap();
    let mut added = Vec::new();
    for slide in 0..count {
        let written = edit
            .add(
                Position::new(slide),
                0,
                0,
                12.0,
                Source::Latex("k"),
                &Limits::DEFAULT,
            )
            .unwrap();
        added.push((slide, written.shape_id()));
    }
    let after = edit.commit().unwrap().into_parts().0;
    let presentation = after.presentation().unwrap();
    for (slide, shape_id) in added {
        let record = presentation
            .persisted_record_bytes(slide_persist(&after, slide))
            .unwrap();
        let mut node = super::tree::parse(record).unwrap();
        node.expand_all(0).unwrap();
        let mut ids = Vec::new();
        super::edit::shape_ids(&node, &mut ids, 0);
        assert_eq!(
            ids.iter().filter(|id| **id == shape_id).count(),
            1,
            "slide {slide}"
        );
        assert!(
            presentation
                .slide_equations(slide)
                .unwrap()
                .iter()
                .any(|equation| equation.shape_id() == shape_id)
        );
    }
}

/// `source` with its one equation shown by `frames` frames (the copies get
/// fresh shape identifiers) and its `Equation Native` stream padded by
/// `padding` zero bytes.
fn shared_heavy_equation(frames: usize, padding: usize) -> Vec<u8> {
    let package = open(&authored(1));
    let mut edit = package.edit_equations().unwrap();
    edit.add(
        Position::new(0),
        0,
        0,
        12.0,
        Source::Latex("x"),
        &Limits::DEFAULT,
    )
    .unwrap();
    let package = edit.commit().unwrap().into_parts().0;
    let equation = equations(&package).remove(0);
    let bytes = package.to_bytes().unwrap();
    let mut editor =
        crate::embedded::object::Editor::open_records(bytes, crate::SourceLimits::default())
            .unwrap();
    let slide_id = slide_persist(&package, 0);
    let mut slide = super::tree::parse(&editor.persisted_record(slide_id).unwrap()).unwrap();
    slide.expand_all(0).unwrap();
    let group = super::edit::patriarch_for_tests(&mut slide);
    let frame = group
        .children()
        .iter()
        .find(|child| read::shape_id(child) == Some(equation.shape_id()))
        .unwrap()
        .clone();
    for copy in 1..frames {
        let mut frame = frame.clone();
        frame.expand_all(0).unwrap();
        let fsp = frame
            .child_mut(0xF00A)
            .unwrap()
            .unwrap()
            .payload_mut()
            .unwrap();
        let id = equation.shape_id() + 100 + u32::try_from(copy).unwrap();
        fsp[..4].copy_from_slice(&id.to_le_bytes());
        group.children_mut().unwrap().push(frame);
    }
    editor
        .replace_persisted_record(slide_id, slide.encode().unwrap())
        .unwrap();
    let mut native = equation.equation_native().unwrap().to_vec();
    native.resize(native.len() + padding, 0);
    let storage =
        super::object::new_storage(&litchi_equation::MtefTarget::MathType, &native).unwrap();
    editor
        .replace_persisted_record(
            equation.persist_id(),
            super::object::storage_record(storage, true).unwrap(),
        )
        .unwrap();
    editor.finish().unwrap()
}

#[test]
fn listing_reads_no_storage_and_each_object_is_read_once() {
    let bytes = shared_heavy_equation(41, 8 * 1024 * 1024);
    let package = open(&bytes);
    let presentation = package.presentation().unwrap();
    let found = presentation.equations().unwrap();
    assert_eq!(found.len(), 41);
    assert_eq!(presentation.slide_equations(0).unwrap().len(), 41);
    assert_eq!(presentation.inner.equation_memo.loads(), 0);
    let first = found[0].equation_native().unwrap();
    assert!(first.len() > 8 * 1024 * 1024);
    for equation in &found {
        assert!(std::sync::Arc::ptr_eq(
            &equation.equation_native().unwrap(),
            &first
        ));
    }
    assert_eq!(presentation.inner.equation_memo.loads(), 1);
}

#[test]
fn reading_equations_stays_within_the_aggregate_payload_limit() {
    let bytes = shared_heavy_equation(3, 8 * 1024 * 1024);
    let limits = crate::SourceLimits::builder()
        .max_aggregate_payload_bytes(4 * 1024 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let package =
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(limits))
            .unwrap();
    let presentation = package.presentation().unwrap();
    let found = presentation.equations().unwrap();
    assert_eq!(found.len(), 3);
    for equation in &found {
        assert!(equation.equation_native().is_err());
        assert!(equation.latex(&Limits::DEFAULT).is_err());
    }
    assert!(presentation.text().is_ok());
}

#[test]
fn a_corrupt_storage_fails_only_its_equation() {
    // Flip bytes inside the zlib stream of the fixture's equation storage
    // (the `ExOleObjStg` at offset 62126 of `PowerPoint Document`).
    let package = open(FIXTURE);
    let document_offset = {
        let mut file = litchi_cfb::OleFile::open(Cursor::new(FIXTURE), Default::default()).unwrap();
        let document = file.open_stream(&["PowerPoint Document"]).unwrap();
        assert_eq!(&document[62126 + 2..62126 + 4], &0x1011u16.to_le_bytes());
        document
    };
    let mut document = document_offset;
    for byte in &mut document[62126 + 8 + 4 + 40..62126 + 8 + 4 + 120] {
        *byte ^= 0x5A;
    }
    let mut editor = crate::embedded::object::Editor::open_records(
        FIXTURE.to_vec(),
        crate::SourceLimits::default(),
    )
    .unwrap();
    let corrupt = document[62126..62126 + 8 + 939].to_vec();
    editor.replace_persisted_record(21, corrupt).unwrap();
    let corrupted = open(&editor.finish().unwrap());
    let presentation = corrupted.presentation().unwrap();
    let found = presentation.equations().unwrap();
    assert_eq!(found.len(), 1);
    assert_eq!(found[0].shape_id(), SHAPE);
    assert!(found[0].equation_native().is_err());
    assert!(matches!(
        found[0].latex(&Limits::DEFAULT),
        Err(litchi_equation::Error::Malformed { .. })
    ));
    assert_eq!(
        presentation.text().unwrap(),
        package.presentation().unwrap().text().unwrap()
    );
    // A replacement needs the old equation and is refused; a removal is not.
    let mut edit = corrupted.edit_equations().unwrap();
    let error = edit
        .set(
            Position::new(SLIDE),
            SHAPE,
            Source::Latex("x"),
            &Limits::DEFAULT,
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::UnsupportedContent),
        "{error:?}"
    );
    edit.remove(Position::new(SLIDE), SHAPE).unwrap();
    assert!(equations(&edit.commit().unwrap().into_parts().0).is_empty());
}

#[test]
fn new_shape_identifiers_are_unused_on_every_slide() {
    let package = open(FIXTURE);
    let count = package.presentation().unwrap().slide_count();
    let mut edit = package.edit_equations().unwrap();
    let mut added = Vec::new();
    for slide in 0..count {
        added.push(
            edit.add(
                Position::new(slide),
                0,
                0,
                12.0,
                Source::Latex("k"),
                &Limits::DEFAULT,
            )
            .unwrap()
            .shape_id(),
        );
    }
    let after = edit.commit().unwrap().into_parts().0;
    let presentation = after.presentation().unwrap();
    let mut everywhere = Vec::new();
    for slide in 0..count {
        let record = presentation
            .persisted_record_bytes(slide_persist(&after, slide))
            .unwrap();
        let mut node = super::tree::parse(record).unwrap();
        node.expand_all(0).unwrap();
        super::edit::shape_ids(&node, &mut everywhere, 0);
    }
    for shape_id in added {
        assert_eq!(
            everywhere.iter().filter(|id| **id == shape_id).count(),
            1,
            "{shape_id}"
        );
    }
}

#[test]
fn an_out_of_range_cluster_table_is_refused() {
    let source = authored(1);
    let package = open(&source);
    let mut editor = crate::embedded::object::Editor::open_records(
        source.clone(),
        crate::SourceLimits::default(),
    )
    .unwrap();
    let document_id = document_persist(&package);
    let mut document = super::tree::parse(&editor.persisted_record(document_id).unwrap()).unwrap();
    let fdgg = super::object::drawing_group(&mut document)
        .unwrap()
        .child_mut(0xF006)
        .unwrap()
        .unwrap()
        .payload_mut()
        .unwrap();
    // The first cluster claims 0xFFFFFFFF used identifiers.
    fdgg[20..24].copy_from_slice(&u32::MAX.to_le_bytes());
    editor
        .replace_persisted_record(document_id, document.encode().unwrap())
        .unwrap();
    let broken = open(&editor.finish().unwrap());
    let mut edit = broken.edit_equations().unwrap();
    let error = edit
        .add(
            Position::new(0),
            0,
            0,
            12.0,
            Source::Latex("x"),
            &Limits::DEFAULT,
        )
        .unwrap_err();
    assert!(
        matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::UnsupportedContent),
        "{error:?}"
    );
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    out.extend_from_slice(&kind.to_le_bytes());
    out.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    out.extend_from_slice(payload);
    out
}

fn words(values: &[u32]) -> Vec<u8> {
    values
        .iter()
        .flat_map(|value| value.to_le_bytes())
        .collect()
}

#[test]
fn references_to_a_frame_are_found_in_timing_build_lists_and_rules() {
    use super::edit::shape_reference;
    // A `VisualShapeAtom` of shape 77 inside a programmable tag's data blob.
    let visual = record(0, 0, 0x2AFB, &words(&[0, 1, 77, 0xFFFF_FFFF, 0xFFFF_FFFF]));
    let time_node = record(0xF, 0, 0xF144, &visual);
    let blob = record(0, 0, 0x138B, &time_node);
    let tag = record(0xF, 0, 0x138A, &blob);
    let slide = record(0xF, 0, 0x03EE, &record(0xF, 0, 0x1388, &tag));
    assert!(shape_reference(&slide, 77).is_some_and(|reason| reason.contains("animation")));
    assert_eq!(shape_reference(&slide, 78), None);
    // A sound (refType 2) with the same id is not a shape.
    let sound = record(0, 0, 0x2AFB, &words(&[0, 2, 77, 0, 0]));
    assert_eq!(shape_reference(&sound, 77), None);
    // A build-list entry and a connector rule.
    let build = record(0, 0, 0x2B03, &[words(&[1, 2, 77]), vec![0; 4]].concat());
    assert!(shape_reference(&build, 77).is_some_and(|reason| reason.contains("build list")));
    let rule = record(0, 0, 0xF012, &words(&[1, 5, 77, 9, 0, 0]));
    let solver = record(0xF, 0, 0xF005, &rule);
    assert!(shape_reference(&solver, 77).is_some_and(|reason| reason.contains("connector")));
    assert_eq!(shape_reference(&solver, 6), None);
}

#[test]
fn a_frame_a_connector_is_attached_to_is_not_removed() {
    let package = open(&authored(1));
    let mut edit = package.edit_equations().unwrap();
    let shape = edit
        .add(
            Position::new(0),
            0,
            0,
            12.0,
            Source::Latex("x"),
            &Limits::DEFAULT,
        )
        .unwrap()
        .shape_id();
    let package = edit.commit().unwrap().into_parts().0;
    let slide_id = slide_persist(&package, 0);
    let mut editor = crate::embedded::object::Editor::open_records(
        package.to_bytes().unwrap(),
        crate::SourceLimits::default(),
    )
    .unwrap();
    let mut slide = super::tree::parse(&editor.persisted_record(slide_id).unwrap()).unwrap();
    slide.expand_all(0).unwrap();
    // An `OfficeArtSolverContainer` after the patriarch, with a connector
    // rule attaching shape `shape` (MS-ODRAW 2.2.18).
    let container = slide
        .child_mut(0x040C)
        .unwrap()
        .unwrap()
        .child_mut(0xF002)
        .unwrap()
        .unwrap();
    let rule = super::tree::Node::atom(
        0,
        0,
        0xF012,
        words(&[1, shape, shape + 50, shape + 51, 0, 0]),
    );
    container
        .children_mut()
        .unwrap()
        .push(super::tree::Node::container(0xF005, 1, vec![rule]));
    editor
        .replace_persisted_record(slide_id, slide.encode().unwrap())
        .unwrap();
    let connected = open(&editor.finish().unwrap());
    let mut edit = connected.edit_equations().unwrap();
    let error = edit.remove(Position::new(0), shape).unwrap_err();
    assert!(
        matches!(&error, Error::Refused(refusal) if refusal.kind() == RefusalKind::UnsupportedContent),
        "{error:?}"
    );
    assert!(error.to_string().contains("connector"), "{error}");
    assert!(!edit.commit().unwrap().diagnostics().changed());
}

#[test]
fn a_large_equation_gets_a_coarser_preview_instead_of_a_refusal() {
    // At 3000 points and the preview's density the picture exceeds the
    // renderer's pixel ceiling.
    let native = litchi_equation::latex_to_equation_native(
        "ab",
        &litchi_equation::MtefTarget::MathType,
        &Limits::DEFAULT,
    )
    .unwrap();
    let dense = litchi_equation::PngOptions::new(3000.0)
        .unwrap()
        .with_scale(2.0)
        .unwrap();
    assert!(matches!(
        litchi_equation::mtef_to_png(native.bytes(), &dense, &Limits::DEFAULT),
        Err(litchi_equation::Error::Render(_))
    ));
    let package = open(&authored(1));
    let mut edit = package.edit_equations().unwrap();
    let written = edit
        .add(
            Position::new(0),
            0,
            0,
            3000.0,
            Source::Latex("ab"),
            &Limits::DEFAULT,
        )
        .unwrap();
    let after = edit.commit().unwrap().into_parts().0;
    let equation = equations(&after)
        .into_iter()
        .find(|equation| equation.shape_id() == written.shape_id())
        .unwrap();
    // The frame keeps the equation's size in points.
    let bounds = equation.bounds().unwrap();
    assert!(bounds.height() > 2000 * 8, "{bounds:?}");
}

#[test]
fn storages_that_fail_to_parse_spend_the_payload_budget() {
    // Sixteen equation objects whose storages decompress to 1 MiB of 0xAB:
    // no compound file, so every read fails after decompressing.
    let package = open(&authored(1));
    let mut edit = package.edit_equations().unwrap();
    for index in 0..16 {
        edit.add(
            Position::new(0),
            index * 10,
            0,
            12.0,
            Source::Latex("x"),
            &Limits::DEFAULT,
        )
        .unwrap();
    }
    let package = edit.commit().unwrap().into_parts().0;
    let mut editor = crate::embedded::object::Editor::open_records(
        package.to_bytes().unwrap(),
        crate::SourceLimits::default(),
    )
    .unwrap();
    for equation in equations(&package) {
        let garbage = super::object::storage_record(vec![0xAB; 1024 * 1024], true).unwrap();
        editor
            .replace_persisted_record(equation.persist_id(), garbage)
            .unwrap();
    }
    let bytes = editor.finish().unwrap();
    let limits = crate::SourceLimits::builder()
        .max_aggregate_payload_bytes(4 * 1024 * 1024)
        .unwrap()
        .build()
        .unwrap();
    let package =
        Package::from_bytes_with_options(&bytes, crate::OpenOptions::new().source_limits(limits))
            .unwrap();
    let presentation = package.presentation().unwrap();
    let found = presentation.equations().unwrap();
    assert_eq!(found.len(), 16);
    let reasons = found
        .iter()
        .map(|equation| equation.equation_native().unwrap_err().reason().to_owned())
        .collect::<Vec<_>>();
    // Four storages fit the 4 MiB budget; the others are not decompressed.
    assert_eq!(presentation.inner.equation_memo.loads(), 4);
    assert!(
        reasons[4..]
            .iter()
            .all(|reason| reason.contains("aggregate payload limit"))
    );
    assert!(!reasons[0].contains("aggregate payload limit"));
}

/// PowerPoint 2007 decks whose `RoundTripCustomTableStyles12Atom` has
/// recVer 0xF while its payload is not a record sequence.
const POWERPOINT_2007: [&[u8]; 3] = [
    include_bytes!("../../../../test-data/poi/test-data/slideshow/45543.ppt"),
    include_bytes!("../../../../test-data/poi/test-data/slideshow/headers_footers_2007.ppt"),
    include_bytes!("../../../../test-data/poi/test-data/slideshow/text-margins.ppt"),
];

#[test]
fn powerpoint_2007_decks_take_equations_and_object_edits() {
    for bytes in POWERPOINT_2007 {
        let package = open(bytes);
        let mut edit = package.edit_equations().unwrap();
        let shape = edit
            .add(
                Position::new(0),
                100,
                100,
                18.0,
                Source::Latex("x+1"),
                &Limits::DEFAULT,
            )
            .unwrap()
            .shape_id();
        let added = edit.commit().unwrap().into_parts().0;
        assert_eq!(latex(&equations(&added)[0]), "x+ 1");
        let mut edit = added.edit_equations().unwrap();
        edit.remove(Position::new(0), shape).unwrap();
        let removed = edit.commit().unwrap().into_parts().0;
        assert!(equations(&removed).is_empty());

        // The embedded-object editor rewrites the same document record.
        let mut objects = package.edit_objects().unwrap();
        let id = objects.objects().unwrap().id_seed + 1;
        let storage = crate::embedded::storage::Storage::uncompressed(
            crate::embedded::storage::Kind::OleObject,
            super::object::new_storage(&litchi_equation::MtefTarget::MathType, b"\x1c\0").unwrap(),
        )
        .unwrap();
        objects
            .add(
                crate::embedded::object::ExternalObject::Object(
                    crate::embedded::object::Definition {
                        kind: crate::embedded::object::ContainerKind::Embedded(
                            crate::embedded::object::EmbedPreferences {
                                color_follow: crate::embedded::object::ColorFollow::None,
                                cannot_lock_server: true,
                                dimension_policy: crate::embedded::object::DimensionPolicy::Send,
                                is_word_table: false,
                                unused: 0,
                            },
                        ),
                        object: crate::embedded::object::Metadata {
                            draw_aspect: crate::embedded::object::DrawAspect::Content,
                            object_type: crate::embedded::object::ObjectType::Embedded,
                            id,
                            subtype: crate::embedded::object::ObjectSubtype::Default,
                            persist_id: 0,
                            unused: [0; 4],
                        },
                        menu_name: None,
                        program_id: None,
                        clipboard_name: None,
                        metafile: None,
                    },
                ),
                storage,
            )
            .unwrap();
        let committed = objects.commit().unwrap().into_parts().0;
        assert!(
            committed
                .presentation()
                .unwrap()
                .ole_objects()
                .unwrap()
                .unwrap()
                .get(id)
                .is_some()
        );
    }
}
