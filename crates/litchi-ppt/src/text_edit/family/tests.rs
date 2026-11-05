//! The source-backed text family through its public route (record 0357a):
//! the scenarios of the native `SourceSnapshot`/`SourceTransaction` tests,
//! run through `Package::{edit_text, apply_text}`.

use std::io::{self, Cursor, Write};
use std::sync::Arc;

use litchi_cfb::{OleFile, OleWriter};
use litchi_core::OwnedSource;

use super::super::{Error, Position, Refusal, Target};
use crate::Package;
use crate::writer::Builder;

fn fixture(texts: &[&str]) -> Vec<u8> {
    let mut writer = Builder::new();
    let slide = writer.add_slide().unwrap();
    for (index, text) in texts.iter().enumerate() {
        let y = 10 + i32::try_from(index).unwrap() * 50;
        writer.add_textbox(slide, 10, y, 240, 40, text).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    crate::raw::encode(&mut writer, &mut output).unwrap();
    output.into_inner()
}

fn with_stream(source: Vec<u8>, path: &[&str], data: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(source), litchi_cfb::OleFileLimits::default()).unwrap();
    let mut writer = OleWriter::new();
    for stream in ole.list_streams() {
        let names: Vec<&str> = stream.iter().map(String::as_str).collect();
        let bytes = ole.open_stream(&names).unwrap();
        writer.create_stream(&names, &bytes).unwrap();
    }
    writer.create_stream(path, data).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn shape(index: usize) -> Target {
    Target::new(Position::new(0), Position::new(index))
}

/// A package over a positional source, as `Package::open` builds one.
fn positional(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn texts(package: &Package) -> Vec<String> {
    package.presentation().unwrap().slides().unwrap()[0]
        .shapes()
        .unwrap()
        .iter()
        .map(|shape| shape.text().unwrap())
        .collect()
}

fn streamed(package: &Package) -> Vec<u8> {
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

#[test]
fn text_edits_publish_replay_and_invert_exactly() {
    let bytes = with_stream(fixture(&["abc", "def"]), &["UnrelatedPayload"], b"opaque");
    for source in [Package::from_bytes(&bytes).unwrap(), positional(&bytes)] {
        let mut edit = source.edit_text(shape(0)).unwrap();
        assert_eq!(edit.text(), "abc");
        assert_eq!(edit.target(), shape(0));
        edit.set_text("xyz").unwrap();
        let commit = edit.commit().unwrap();
        let diagnostics = commit.diagnostics();
        assert!(diagnostics.changed());
        assert_eq!(diagnostics.target(), shape(0));
        assert_eq!(diagnostics.replacement_bytes(), 3);
        assert_eq!(diagnostics.source_bytes(), bytes.len() as u64);
        assert_ne!(
            diagnostics.source_fingerprint(),
            diagnostics.target_fingerprint()
        );
        assert!(!commit.patch().is_noop());
        assert_eq!(commit.patch().target(), shape(0));

        // The published package reads the splice; the other shape, the
        // other streams and the length are unchanged.
        let published = commit.package();
        assert_eq!(texts(published), ["xyz", "def"]);
        assert_eq!(published.edit_text(shape(0)).unwrap().text(), "xyz");
        let output = streamed(published);
        assert_eq!(output, published.to_bytes().unwrap());
        assert_eq!(output.len(), bytes.len());
        let mut original = OleFile::open(
            Cursor::new(bytes.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        let mut rewritten = OleFile::open(
            Cursor::new(output.clone()),
            litchi_cfb::OleFileLimits::default(),
        )
        .unwrap();
        assert_eq!(
            rewritten.open_stream(&["UnrelatedPayload"]).unwrap(),
            original.open_stream(&["UnrelatedPayload"]).unwrap()
        );
        let differing = bytes
            .iter()
            .zip(&output)
            .filter(|(left, right)| left != right)
            .count();
        assert!((1..=3).contains(&differing));

        // Replay on the source (and on an equal package), inverse on the
        // result.
        let replayed = source.apply_text(commit.patch()).unwrap();
        assert!(replayed.diagnostics().changed());
        assert_eq!(streamed(replayed.package()), output);
        let reopened = positional(&bytes).apply_text(commit.patch()).unwrap();
        assert_eq!(streamed(reopened.package()), output);
        let restored = published.apply_text(&commit.patch().inverse()).unwrap();
        assert_eq!(streamed(restored.package()), bytes);
        assert_eq!(texts(restored.package()), ["abc", "def"]);
        assert_eq!(
            commit.patch().inverse().inverse().source_fingerprint(),
            commit.patch().source_fingerprint()
        );

        // Any other package is a conflict.
        assert!(matches!(
            published.apply_text(commit.patch()),
            Err(Error::PatchConflict)
        ));
        assert!(matches!(
            Package::from_bytes(&fixture(&["abd", "def"]))
                .unwrap()
                .apply_text(commit.patch()),
            Err(Error::PatchConflict)
        ));
    }
}

#[test]
fn chained_text_edits_read_through_the_composed_source() {
    let bytes = fixture(&["abc", "def"]);
    let source = Package::from_bytes(&bytes).unwrap();
    let mut first = source.edit_text(shape(0)).unwrap();
    first.set_text("xyz").unwrap();
    let first = first.commit().unwrap();
    let mut second = first.package().edit_text(shape(1)).unwrap();
    second.set_text("uvw").unwrap();
    let second = second.commit().unwrap();
    assert_eq!(texts(second.package()), ["xyz", "uvw"]);
    let back = second
        .package()
        .apply_text(&second.patch().inverse())
        .unwrap()
        .into_package()
        .apply_text(&first.patch().inverse())
        .unwrap();
    assert_eq!(streamed(back.package()), bytes);
}

#[test]
fn noop_text_commits_return_the_source_package() {
    let bytes = fixture(&["abc"]);
    let source = positional(&bytes);
    for staged in [None, Some("abc")] {
        let mut edit = source.edit_text(shape(0)).unwrap();
        if let Some(text) = staged {
            edit.set_text(text).unwrap();
        }
        let commit = edit.commit().unwrap();
        assert!(!commit.diagnostics().changed());
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.patch().source_fingerprint(),
            commit.patch().target_fingerprint()
        );
        assert_eq!(streamed(commit.package()), bytes);
        let (package, patch) = commit.into_parts();
        let applied = package.apply_text(&patch).unwrap();
        assert!(!applied.diagnostics().changed());
        assert!(matches!(
            Package::from_bytes(&fixture(&["abd"]))
                .unwrap()
                .apply_text(&patch),
            Err(Error::PatchConflict)
        ));
    }
}

#[test]
fn text_edits_refuse_length_changes_and_absent_targets() {
    let source = Package::from_bytes(&fixture(&["abc"])).unwrap();
    let mut edit = source.edit_text(shape(0)).unwrap();
    assert!(matches!(
        edit.set_text("longer"),
        Err(Error::Refused(Refusal::LengthChange))
    ));
    assert_eq!(edit.text(), "abc");
    assert!(!edit.commit().unwrap().diagnostics().changed());

    let missing_slide = source.edit_text(Target::new(Position::new(1), Position::new(0)));
    assert!(matches!(
        missing_slide,
        Err(Error::Refused(refusal)) if refusal.kind() == litchi_core::edit::RefusalKind::NotFound
    ));
    assert!(matches!(
        source.edit_text(shape(1)),
        Err(Error::Refused(Refusal::ShapeNotFound))
    ));
}

#[test]
fn text_edits_refuse_unsupported_sources() {
    let bytes = with_stream(fixture(&["abc"]), &["VBA", "dir"], b"macro");
    let source = Package::from_bytes(&bytes).unwrap();
    assert!(matches!(
        source.edit_text(shape(0)),
        Err(Error::Refused(Refusal::UnsupportedSource))
    ));
}

#[test]
fn signed_packages_admit_noop_text_commits_only() {
    let bytes = with_stream(fixture(&["abc"]), &["_signatures"], b"opaque signature");
    let source = Package::from_bytes(&bytes).unwrap();
    let mut edit = source.edit_text(shape(0)).unwrap();
    assert_eq!(edit.text(), "abc");
    edit.set_text("abc").unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(streamed(commit.package()), bytes);
    assert!(source.apply_text(commit.patch()).is_ok());

    let mut edit = source.edit_text(shape(0)).unwrap();
    edit.set_text("xyz").unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));

    // A changed patch made on the unsigned twin is refused as signed only
    // when the fingerprints match; here they differ, so it is a conflict.
    let unsigned = Package::from_bytes(&fixture(&["abc"])).unwrap();
    let mut edit = unsigned.edit_text(shape(0)).unwrap();
    edit.set_text("xyz").unwrap();
    let patch = edit.commit().unwrap().patch().clone();
    assert!(matches!(
        source.apply_text(&patch),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn published_text_package_streams_and_reports_partial_sinks() {
    struct PartialSink {
        bytes: Vec<u8>,
        limit: usize,
    }
    impl Write for PartialSink {
        fn write(&mut self, data: &[u8]) -> io::Result<usize> {
            let room = self.limit.saturating_sub(self.bytes.len());
            if room == 0 {
                return Err(io::Error::other("sink full"));
            }
            let count = room.min(data.len());
            self.bytes.extend_from_slice(&data[..count]);
            Ok(count)
        }
        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }

    let source = positional(&fixture(&["abc"]));
    let mut edit = source.edit_text(shape(0)).unwrap();
    edit.set_text("xyz").unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = PartialSink {
        bytes: Vec::new(),
        limit: 11,
    };
    assert!(commit.package().write_to(&mut sink).is_err());
    assert_eq!(sink.bytes.len(), 11);
}
