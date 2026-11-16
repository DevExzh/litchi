//! Equation objects inserted, replaced and removed through the primary DOC
//! package transaction, on Word-written equations of both kinds and on
//! documents written by litchi. Every edit is read back through litchi and
//! through the MTEF reader, its patch inverts to the source bytes, and the
//! streams it does not own keep their bytes.
#![cfg(feature = "equation")]

use litchi_cfb::{OleFile, OleFileLimits};
use litchi_core::Position;
use litchi_doc::body_text::{Error as BodyError, Refusal};
use litchi_doc::builder::Builder;
use litchi_doc::{EquationContent, EquationKind, Package};
use litchi_equation::{Limits, MtefTarget};
use std::collections::BTreeMap;
use std::io::Cursor;

/// 671 equations: 327 Equation Editor 3 and 344 MathType objects.
const TRIGONOMETRIC_RATIOS: &[u8] =
    include_bytes!("../../../test-data/ole/doc/trigonometric_ratios.doc");

/// Every stream of a compound file, by path.
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec()), OleFileLimits::default()).unwrap();
    let paths = ole.list_streams();
    paths
        .into_iter()
        .map(|path| {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&refs).unwrap();
            (path, data)
        })
        .collect()
}

/// The equations of a document as (field begin CP, storage number, LaTeX),
/// in document order, read through litchi's reader.
fn equations(bytes: &[u8]) -> Vec<(u32, u32, String)> {
    let mut reader = litchi_doc::raw::Reader::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    let document = reader.document().unwrap();
    let mut objects = Vec::new();
    for paragraph in document.paragraphs().unwrap() {
        for run in paragraph.runs().unwrap() {
            if let Some(latex) = run.mtef_equation_latex() {
                objects.push((run.properties().pic_offset.unwrap(), latex.to_owned()));
            }
        }
    }
    let fields: Vec<u32> = document
        .fields()
        .unwrap()
        .iter()
        .filter(|field| field.instruction.trim_start().starts_with("EMBED Equation"))
        .map(|field| field.field.start_cp)
        .collect();
    assert_eq!(fields.len(), objects.len());
    fields
        .into_iter()
        .zip(objects)
        .map(|(start, (storage, latex))| (start, storage, latex))
        .collect()
}

fn equation_native(bytes: &[u8], storage: u32) -> Option<Vec<u8>> {
    streams(bytes).remove(&vec![
        "ObjectPool".to_owned(),
        format!("_{storage}"),
        "Equation Native".to_owned(),
    ])
}

fn storage_clsid(bytes: &[u8], storage: u32) -> Option<String> {
    let ole = OleFile::open(Cursor::new(bytes.to_vec()), OleFileLimits::default()).unwrap();
    ole.list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == format!("_{storage}"))
        .map(|entry| entry.clsid.clone())
}

/// Commit `edit`, check that the patch inverts to the exact source bytes,
/// and return the published bytes.
fn commit(source: &Package, edit: litchi_doc::Edit<'static>) -> Vec<u8> {
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let (package, patch) = commit.into_parts();
    let restored = package.apply(&patch.inverse()).unwrap();
    assert_eq!(
        restored.package().to_bytes().unwrap(),
        source.to_bytes().unwrap()
    );
    let bytes = package.to_bytes().unwrap();
    if let Some(directory) = std::env::var_os("LITCHI_DOC_EQUATION_OUT") {
        static COMMITS: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);
        let commit = COMMITS.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let thread = std::thread::current();
        let name = format!("{}-{commit}.doc", thread.name().unwrap_or("out"));
        std::fs::write(
            std::path::Path::new(&directory).join(name.replace("::", "-")),
            &bytes,
        )
        .unwrap();
    }
    bytes
}

fn refusal(error: litchi_doc::package::edit::Error) -> Refusal {
    match error {
        litchi_doc::package::edit::Error::Body(BodyError::Refused(refusal)) => refusal,
        other => panic!("expected a refusal, got {other:?}"),
    }
}

fn latex_content(latex: &str, target: &MtefTarget) -> EquationContent {
    let (content, losses) =
        EquationContent::from_latex(latex, target, 12.0, &Limits::DEFAULT).unwrap();
    assert!(losses.is_empty(), "{losses:?}");
    content
}

fn mtef_latex(bytes: &[u8]) -> String {
    litchi_equation::mtef_to_latex(bytes, &Limits::DEFAULT).unwrap()
}

fn builder_doc() -> Vec<u8> {
    let mut writer = Builder::new();
    writer.add_paragraph("Before the equation.").unwrap();
    writer.add_paragraph("The area is  square units.").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

/// The streams of `after` that differ from `before`, and those only one of
/// them has.
fn changed_streams(before: &[u8], after: &[u8]) -> Vec<Vec<String>> {
    let (before, after) = (streams(before), streams(after));
    let mut changed: Vec<Vec<String>> = before
        .iter()
        .filter(|(path, data)| after.get(*path) != Some(*data))
        .map(|(path, _)| path.clone())
        .collect();
    changed.extend(
        after
            .keys()
            .filter(|path| !before.contains_key(*path))
            .cloned(),
    );
    changed.sort();
    changed
}

/// The `sprmCHpsPos` of the picture character that shows the preview of
/// the object in storage `storage`: the character after its field
/// separator, in half-points (negative is lowered).
fn picture_position(bytes: &[u8], storage: u32) -> i16 {
    let mut reader = litchi_doc::raw::Reader::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    let document = reader.document().unwrap();
    let mut runs = Vec::new();
    for paragraph in document.paragraphs().unwrap() {
        runs.extend(paragraph.runs().unwrap());
    }
    let separator = runs
        .iter()
        .position(|run| run.properties().is_ole2 && run.properties().pic_offset == Some(storage))
        .unwrap();
    let picture = &runs[separator + 1];
    assert_eq!(picture.text().unwrap(), "\u{1}");
    picture.properties().position.half_points()
}

/// The length of the main-story field table (`PlcfFldMom`) in the FIB.
fn field_table_length(bytes: &[u8]) -> u32 {
    let word = &streams(bytes)[&path(&["WordDocument"])];
    let at = 154 + 16 * 8 + 4;
    u32::from_le_bytes(word[at..at + 4].try_into().unwrap())
}

fn path(parts: &[&str]) -> Vec<String> {
    parts.iter().map(|part| (*part).to_owned()).collect()
}

#[test]
fn a_mathtype_equation_is_inserted_inline_in_a_litchi_document() {
    let bytes = builder_doc();
    let source = Package::from_bytes(&bytes).unwrap();
    let content = latex_content(r"A=\pi r^{2}", &MtefTarget::MathType);
    let mut edit = source.edit().unwrap();
    // Between the two spaces of "The area is  square units.".
    let storage = edit
        .insert_equation(Position::new(1), 12, EquationKind::MathType, &content)
        .unwrap();
    let after = commit(&source, edit);

    let found = equations(&after);
    assert_eq!(found.len(), 1);
    let (_, shown, latex) = &found[0];
    assert_eq!(*shown, storage);
    assert_eq!(latex, r"A= \pi r^{2}");
    // The preview is lowered so that the equation's baseline sits on the
    // text line, as Word lowers every equation.
    assert_eq!(
        picture_position(&after, storage),
        -i16::try_from(content.lowering()).unwrap()
    );
    assert!(content.lowering() > 0);
    let native = equation_native(&after, storage).unwrap();
    assert_eq!(native, content.equation().equation_native());
    assert_eq!(native[28], 5);
    assert_eq!(mtef_latex(&native), *latex);
    assert_eq!(
        storage_clsid(&after, storage).as_deref(),
        Some("0002CE03-0000-0000-C000-000000000046")
    );
    let all = streams(&after);
    let comp_obj = &all[&path(&["ObjectPool", &format!("_{storage}"), "\u{1}CompObj"])];
    assert!(
        comp_obj
            .windows(b"Equation.DSMT4\0".len())
            .any(|window| window == b"Equation.DSMT4\0")
    );
    for stream in ["\u{1}Ole", "\u{3}ObjInfo"] {
        assert!(all.contains_key(&path(&["ObjectPool", &format!("_{storage}"), stream])));
    }

    // The paragraph keeps its text around the field.
    let package = Package::from_bytes(&after).unwrap();
    let text = package
        .document()
        .body_paragraph_text(Position::new(1))
        .unwrap()
        .unwrap()
        .to_owned();
    assert!(text.starts_with("The area is \u{13} EMBED Equation.DSMT4  \u{14}\u{1}\u{15} square"));
    assert_eq!(
        package
            .document()
            .body_paragraph_text(Position::new(0))
            .unwrap(),
        Some("Before the equation.")
    );
    // Only the text, table and Data streams change, and the object is added.
    let mut expected = vec![
        path(&["0Table"]),
        path(&["1Table"]),
        path(&["Data"]),
        path(&["WordDocument"]),
    ];
    expected.extend(
        [
            "\u{1}CompObj",
            "\u{1}Ole",
            "\u{3}ObjInfo",
            "Equation Native",
        ]
        .map(|stream| path(&["ObjectPool", &format!("_{storage}"), stream])),
    );
    let changed = changed_streams(&bytes, &after);
    for changed_path in &changed {
        assert!(expected.contains(changed_path), "{changed_path:?} changed");
    }

    // The new object is updated and removed like a Word-written one.
    let start = found[0].0;
    let source = Package::from_bytes(&after).unwrap();
    let mut edit = source.edit().unwrap();
    let fraction = latex_content(r"\frac{a}{b}", &MtefTarget::MathType);
    // A fraction reaches further below the baseline than `A=πr²`.
    assert!(fraction.lowering() > content.lowering());
    // The preview is drawn at the text's size, and lowered to match.
    let (large, _) = EquationContent::from_latex(
        r"\frac{a}{b}",
        &MtefTarget::MathType,
        24.0,
        &Limits::DEFAULT,
    )
    .unwrap();
    assert!(large.lowering() > fraction.lowering() * 3 / 2);
    edit.set_equation(start, &fraction).unwrap();
    let updated = commit(&source, edit);
    assert_eq!(equations(&updated)[0].2, r"\frac{a}{b}");
    assert_eq!(
        picture_position(&updated, storage),
        -i16::try_from(fraction.lowering()).unwrap()
    );
    let source = Package::from_bytes(&updated).unwrap();
    let mut edit = source.edit().unwrap();
    edit.remove_equation(start).unwrap();
    let removed = commit(&source, edit);
    assert!(equations(&removed).is_empty());
    // The document had no other field: the empty field table is dropped.
    assert_ne!(field_table_length(&updated), 0);
    assert_eq!(field_table_length(&removed), 0);
    assert!(equation_native(&removed, storage).is_none());
    assert_eq!(
        Package::from_bytes(&removed)
            .unwrap()
            .document()
            .body_paragraph_text(Position::new(1))
            .unwrap(),
        Some("The area is  square units.")
    );
}

#[test]
fn equations_of_both_kinds_are_inserted_into_a_word_document() {
    let source = Package::from_bytes(TRIGONOMETRIC_RATIOS).unwrap();
    let before = equations(TRIGONOMETRIC_RATIOS);
    let mut edit = source.edit().unwrap();
    let math_type = latex_content(r"\sin^{2}\alpha+\cos^{2}\alpha=1", &MtefTarget::MathType);
    let editor = latex_content(r"\tan\alpha", &MtefTarget::EquationEditor3);
    let first = edit
        .insert_equation(Position::new(2), 0, EquationKind::MathType, &math_type)
        .unwrap();
    let second = edit
        .insert_equation(Position::new(2), 0, EquationKind::EquationEditor3, &editor)
        .unwrap();
    assert_ne!(first, second);
    let after = commit(&source, edit);

    let found = equations(&after);
    assert_eq!(found.len(), before.len() + 2);
    let new: Vec<_> = found
        .iter()
        .filter(|(_, storage, _)| [first, second].contains(storage))
        .collect();
    assert_eq!(new.len(), 2);
    // The later insertion at the same offset comes first.
    assert_eq!(new[0].1, second);
    assert_eq!(equation_native(&after, second).unwrap()[28], 3);
    assert_eq!(
        storage_clsid(&after, second).as_deref(),
        Some("0002CE02-0000-0000-C000-000000000046")
    );
    assert_eq!(
        mtef_latex(&equation_native(&after, first).unwrap()),
        new[1].2
    );
    // Every Word-written equation keeps its object and its LaTeX.
    let old: Vec<_> = found
        .iter()
        .filter(|(_, storage, _)| ![first, second].contains(storage))
        .map(|(_, storage, latex)| (*storage, latex.clone()))
        .collect();
    let original: Vec<_> = before
        .iter()
        .map(|(_, storage, latex)| (*storage, latex.clone()))
        .collect();
    assert_eq!(old, original);
    for changed in changed_streams(TRIGONOMETRIC_RATIOS, &after) {
        assert!(
            ["1Table", "Data", "WordDocument"].contains(&changed[0].as_str())
                || changed[..2] == path(&["ObjectPool", &format!("_{first}")])
                || changed[..2] == path(&["ObjectPool", &format!("_{second}")]),
            "{changed:?} changed"
        );
    }
}

#[test]
fn word_equations_of_both_kinds_are_updated_in_their_own_kind() {
    let source = Package::from_bytes(TRIGONOMETRIC_RATIOS).unwrap();
    let before = equations(TRIGONOMETRIC_RATIOS);
    // The first equation is Equation Editor 3, the second MathType.
    let (editor_start, editor_storage, _) = before[0].clone();
    let (math_start, math_storage, _) = before[1].clone();
    let info = |storage| {
        litchi_equation::read_mtef(
            &equation_native(TRIGONOMETRIC_RATIOS, storage).unwrap(),
            &Limits::DEFAULT,
        )
        .unwrap()
        .info()
        .cloned()
        .unwrap()
    };
    let editor_content = latex_content(r"x^{2}+1", &MtefTarget::SameAs(info(editor_storage)));
    let math_content = latex_content(r"\sqrt{2}", &MtefTarget::SameAs(info(math_storage)));
    let mut edit = source.edit().unwrap();
    edit.set_equation(editor_start, &editor_content).unwrap();
    edit.set_equation(math_start, &math_content).unwrap();
    let after = commit(&source, edit);

    let found = equations(&after);
    assert_eq!(found.len(), before.len());
    assert_eq!(found[0].2, "x^{2}+ 1");
    assert_eq!(found[1].2, r"\sqrt{2}");
    for (index, (expected, actual)) in before.iter().zip(&found).enumerate().skip(2) {
        assert_eq!(expected, actual, "equation {index}");
    }
    let editor_native = equation_native(&after, editor_storage).unwrap();
    assert_eq!(editor_native[28], 3);
    assert_eq!(mtef_latex(&editor_native), found[0].2);
    assert_eq!(equation_native(&after, math_storage).unwrap()[28], 5);
    // The objects keep their class and every other stream.
    assert_eq!(
        storage_clsid(&after, editor_storage),
        storage_clsid(TRIGONOMETRIC_RATIOS, editor_storage)
    );
    let changed = changed_streams(TRIGONOMETRIC_RATIOS, &after);
    let mut expected = vec![
        path(&["1Table"]),
        path(&["Data"]),
        path(&[
            "ObjectPool",
            &format!("_{editor_storage}"),
            "Equation Native",
        ]),
        path(&["ObjectPool", &format!("_{math_storage}"), "Equation Native"]),
        path(&["WordDocument"]),
    ];
    expected.sort();
    assert_eq!(changed, expected);

    // An equation keeps its kind: MTEF 5 does not replace MTEF 3.
    let source = Package::from_bytes(TRIGONOMETRIC_RATIOS).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        refusal(edit.set_equation(editor_start, &math_content).unwrap_err()),
        Refusal::InvalidProperty
    );
}

#[test]
fn word_equations_are_removed_with_their_storage_and_preview() {
    let source = Package::from_bytes(TRIGONOMETRIC_RATIOS).unwrap();
    let before = equations(TRIGONOMETRIC_RATIOS);
    let (start, storage, _) = before[3].clone();
    let mut edit = source.edit().unwrap();
    edit.remove_equation(start).unwrap();
    let after = commit(&source, edit);

    let found = equations(&after);
    assert_eq!(found.len(), before.len() - 1);
    let mut expected: Vec<_> = before.iter().map(|(_, s, l)| (*s, l.clone())).collect();
    expected.remove(3);
    let actual: Vec<_> = found.iter().map(|(_, s, l)| (*s, l.clone())).collect();
    assert_eq!(actual, expected);
    assert!(equation_native(&after, storage).is_none());
    assert!(storage_clsid(&after, storage).is_none());
    for changed in changed_streams(TRIGONOMETRIC_RATIOS, &after) {
        assert!(
            ["1Table", "Data", "WordDocument"].contains(&changed[0].as_str())
                || changed[..2] == path(&["ObjectPool", &format!("_{storage}")]),
            "{changed:?} changed"
        );
    }
    // The preview's Data bytes are cleared in place: no other offset moves.
    let (old_data, new_data) = (
        &streams(TRIGONOMETRIC_RATIOS)[&path(&["Data"])],
        &streams(&after)[&path(&["Data"])],
    );
    assert_eq!(old_data.len(), new_data.len());
    assert!(
        old_data
            .iter()
            .zip(new_data)
            .filter(|(a, b)| a != b)
            .count()
            > 0
    );

    // A field that is not an equation is refused.
    let source = Package::from_bytes(TRIGONOMETRIC_RATIOS).unwrap();
    let mut edit = source.edit().unwrap();
    assert_eq!(
        refusal(edit.remove_equation(start + 1).unwrap_err()),
        Refusal::EquationNotFound
    );
}
