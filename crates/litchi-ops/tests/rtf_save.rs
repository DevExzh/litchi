#![cfg(feature = "rtf")]

use litchi_core::{Durability, FileFormat, Position};
use litchi_ops::{Code, Overwrite, SaveOptions, SaveTo, create, open, open_bytes};

#[test]
fn native_and_unified_publication_agree_at_every_durability() {
    let dir = tempfile::tempdir().unwrap();
    let package = litchi_rtf::Builder::new()
        .text("before\\{}\n—é")
        .unwrap()
        .build()
        .unwrap();
    let expected = package.to_bytes().unwrap();
    for durability in [Durability::Full, Durability::FileOnly, Durability::NoSync] {
        let path = dir.path().join("native.rtf");
        std::fs::write(&path, b"old").unwrap();
        let count = package.save_with_durability(&path, durability).unwrap();
        assert_eq!(count as usize, expected.len());
        assert_eq!(std::fs::read(&path).unwrap(), expected);
        let file = open(&path).unwrap();
        let output = dir.path().join("unified.rtf");
        let saved = file
            .save_with(
                SaveTo::Path(&output),
                &SaveOptions::new().durability(durability),
            )
            .unwrap();
        assert_eq!(saved.bytes, count);
        assert_eq!(std::fs::read(&output).unwrap(), expected);
        let mut streamed = Vec::new();
        assert_eq!(file.write_to(&mut streamed).unwrap(), count);
        assert_eq!(streamed, expected);
        assert_eq!(open(&output).unwrap().text().unwrap(), package.text());
    }
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 2);
}

#[test]
fn origin_and_explicit_destination_have_distinct_lifetimes() {
    let dir = tempfile::tempdir().unwrap();
    let created = create(FileFormat::Rtf).unwrap();
    assert_eq!(created.save().unwrap_err().code(), Code::FileRequired);
    let origin = dir.path().join("original.rtf");
    created.save_as(&origin).unwrap();
    assert_eq!(created.save().unwrap_err().code(), Code::FileRequired);
    let eager = litchi_ops::open_with(
        &origin,
        &litchi_ops::OpenOptions::new().source_mode(litchi_ops::SourceMode::Eager),
    )
    .unwrap();
    eager.save().unwrap();
    let file = open(&origin).unwrap();
    let copy = dir.path().join("copy.rtf");
    file.save_as(&copy).unwrap();
    std::fs::write(&origin, b"old").unwrap();
    file.save().unwrap();
    assert_eq!(
        std::fs::read(&origin).unwrap(),
        std::fs::read(&copy).unwrap()
    );
    let bytes = std::fs::read(&origin).unwrap();
    assert_eq!(
        open_bytes(bytes.as_slice())
            .unwrap()
            .save()
            .unwrap_err()
            .code(),
        Code::FileRequired
    );
}

#[test]
fn refusal_preserves_destination_and_typed_error() {
    let dir = tempfile::tempdir().unwrap();
    let file = create(FileFormat::Rtf).unwrap();
    let path = dir.path().join("existing.rtf");
    std::fs::write(&path, b"sentinel").unwrap();
    let error = file
        .save_with(
            SaveTo::Path(&path),
            &SaveOptions::new().overwrite(Overwrite::Refuse),
        )
        .unwrap_err();
    assert_eq!(error.format(), Some(FileFormat::Rtf));
    match error.leaf() {
        litchi_ops::error::Source::Publication(litchi_core::atomic::Error::Io(error)) => {
            assert_eq!(error.kind(), std::io::ErrorKind::AlreadyExists)
        },
        other => panic!("unexpected source {other:?}"),
    }
    assert_eq!(std::fs::read(&path).unwrap(), b"sentinel");
    assert!(file.save_as(dir.path().join("wrong.docx")).is_err());
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 1);
    let fresh = dir.path().join("fresh.rtf");
    file.save_with(
        SaveTo::Path(&fresh),
        &SaveOptions::new().overwrite(Overwrite::Refuse),
    )
    .unwrap();
    assert_eq!(open(&fresh).unwrap().text().unwrap(), "");
    assert_eq!(std::fs::read_dir(dir.path()).unwrap().count(), 2);
}

#[cfg(unix)]
#[test]
fn symlinks_are_refused_and_regular_permissions_survive() {
    use std::os::unix::fs::{PermissionsExt, symlink};
    let dir = tempfile::tempdir().unwrap();
    let path = dir.path().join("target.rtf");
    std::fs::write(&path, b"sentinel").unwrap();
    std::fs::set_permissions(&path, std::fs::Permissions::from_mode(0o640)).unwrap();
    let link = dir.path().join("link.rtf");
    symlink(&path, &link).unwrap();
    let package = litchi_rtf::Package::new().unwrap();
    assert!(package.save(&link).is_err());
    assert_eq!(std::fs::read(&path).unwrap(), b"sentinel");
    package.save(&path).unwrap();
    assert_eq!(
        std::fs::metadata(&path).unwrap().permissions().mode() & 0o777,
        0o640
    );
    assert!(
        std::fs::symlink_metadata(&link)
            .unwrap()
            .file_type()
            .is_symlink()
    );
}

#[test]
fn failed_stream_keeps_io_kind_and_rtf_context() {
    struct Failing;
    impl std::io::Write for Failing {
        fn write(&mut self, _: &[u8]) -> std::io::Result<usize> {
            Err(std::io::Error::new(
                std::io::ErrorKind::BrokenPipe,
                "injected",
            ))
        }
        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    let file = create(FileFormat::Rtf).unwrap();
    let error = file.write_to(Failing).unwrap_err();
    assert_eq!(error.format(), Some(FileFormat::Rtf));
    assert!(
        matches!(error.leaf(), litchi_ops::error::Source::Io(source) if source.kind() == std::io::ErrorKind::BrokenPipe)
    );
}

#[test]
fn committed_publication_failure_remains_distinguishable() {
    let error = litchi_ops::Error::from(litchi_core::atomic::Error::Committed {
        source: std::io::Error::other("injected directory sync"),
    });
    assert_eq!(error.code(), Code::InternalError);
    assert!(matches!(
        error.leaf(),
        litchi_ops::error::Source::Publication(litchi_core::atomic::Error::Committed { .. })
    ));
    assert!(std::error::Error::source(&error).is_some());
}

#[test]
fn native_edit_and_retained_navigation_bytes_publish_without_rewriting() {
    let dir = tempfile::tempdir().unwrap();
    for bytes in [
        include_bytes!("../../litchi-rtf/tests/fixtures/navigation-toc.rtf").as_slice(),
        include_bytes!("../../litchi-rtf/tests/fixtures/navigation-index.rtf").as_slice(),
    ] {
        let file = open_bytes(bytes).unwrap();
        let path = dir.path().join("navigation.rtf");
        file.save_as(&path).unwrap();
        assert_eq!(std::fs::read(&path).unwrap(), bytes);
    }
    let original = litchi_rtf::Builder::new()
        .text("before")
        .unwrap()
        .build()
        .unwrap();
    let mut edit = original.edit();
    edit.replace_paragraph_text(Position::new(0), "after")
        .unwrap();
    let changed = edit.commit().unwrap().into_package();
    let path = dir.path().join("changed.rtf");
    changed.save(&path).unwrap();
    assert_eq!(open(&path).unwrap().text().unwrap(), "after");
    assert_eq!(original.text(), "before");
}
