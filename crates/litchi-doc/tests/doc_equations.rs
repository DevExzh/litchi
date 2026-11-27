//! Embedded MathType and Equation Editor equations read as LaTeX.
#![cfg(feature = "equation")]

use litchi_cfb::{OleFile, OleFileLimits};
use litchi_doc::raw::Reader;
use std::collections::HashMap;
use std::io::Cursor;

/// 671 equations: 327 Equation Editor 3 (MTEF 3) objects and 344 MathType
/// (MTEF 5) objects. See `trigonometric_ratios.PROVENANCE.md`.
const TRIGONOMETRIC_RATIOS: &[u8] =
    include_bytes!("../../../test-data/ole/doc/trigonometric_ratios.doc");

/// The `Equation Native` stream of every `ObjectPool` object, read straight
/// from the compound file and keyed by the object's picture location.
fn equation_native_streams(bytes: &[u8]) -> HashMap<u32, Vec<u8>> {
    let mut ole = OleFile::open(Cursor::new(bytes.to_vec()), OleFileLimits::default()).unwrap();
    let storages: Vec<String> = ole
        .list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .map(|entry| entry.name.clone())
        .collect();
    let mut streams = HashMap::new();
    for storage in storages {
        if let Ok(stream) = ole.open_stream(&["ObjectPool", &storage, "Equation Native"]) {
            let id = storage.strip_prefix('_').unwrap().parse().unwrap();
            streams.insert(id, stream);
        }
    }
    streams
}

#[test]
fn every_mtef_3_and_mtef_5_equation_is_read_by_its_picture_location() {
    let streams = equation_native_streams(TRIGONOMETRIC_RATIOS);
    assert_eq!(streams.len(), 671);
    // The MTEF version follows the 28-byte `Equation Native` header.
    let version = |id: &u32| streams[id][28];
    assert_eq!(streams.keys().filter(|id| version(id) == 3).count(), 327);
    assert_eq!(streams.keys().filter(|id| version(id) == 5).count(), 344);

    let mut package = Reader::from_reader(Cursor::new(TRIGONOMETRIC_RATIOS.to_vec())).unwrap();
    let document = package.document().unwrap();
    let mut equations = Vec::new();
    for paragraph in document.paragraphs().unwrap() {
        let runs = paragraph.runs().unwrap();
        assert_eq!(
            paragraph.has_equations(),
            runs.iter().any(|run| run.has_mtef_equation())
        );
        for run in runs {
            if let Some(latex) = run.mtef_equation_latex() {
                let id = run.properties().pic_offset.unwrap();
                equations.push((id, latex.to_owned()));
            }
        }
    }

    // Each object is shown by exactly one run, which carries that object's
    // own equation.
    assert_eq!(equations.len(), 671);
    let mut shown: Vec<u32> = equations.iter().map(|(id, _)| *id).collect();
    shown.sort_unstable();
    shown.dedup();
    assert_eq!(shown.len(), 671);
    for (id, latex) in &equations {
        assert!(!latex.is_empty(), "equation _{id} has no LaTeX");
        let own = litchi_equation::mtef_to_latex(&streams[id], &litchi_equation::Limits::DEFAULT)
            .unwrap();
        assert_eq!(latex, &own, "equation _{id}");
    }
    assert!(
        !document
            .tolerance_report()
            .defects()
            .iter()
            .any(|defect| matches!(defect.defect, litchi_doc::Defect::Equation(_)))
    );

    // The first equations of the document, as printed on its first two
    // pages, with the MTEF version of each.
    let expected = [
        (0, 3, r"P(r\cos \alpha,r\sin \alpha)"),
        (1, 5, r"\alpha= \frac{l}{r}"),
        (2, 5, r"S= \frac{1}{2}rl= \frac{1}{2}\alpha r^{2}"),
        (3, 3, r"0< \alpha< 2\pi"),
        (7, 5, r"180\degree= \pi"),
        (11, 5, r"(2k\pi+ \frac{3\pi}{2},2k\pi+ 2\pi)"),
        (25, 5, r"2k\pi\pm \frac{\pi}{2}"),
        (26, 3, r"\frac{k\pi}{2}"),
        (31, 5, r"\left| OP \right|= r= \sqrt{x^{2}+ y^{2}}"),
    ];
    for (index, mtef_version, latex) in expected {
        let (id, actual) = &equations[index];
        assert_eq!(version(id), mtef_version, "equation {index}");
        assert_eq!(actual, latex, "equation {index}");
    }
}
