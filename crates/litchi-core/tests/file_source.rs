#![cfg(any(unix, windows))]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic by design"
)]

use std::{
    fs::{self, File, OpenOptions},
    io::{self, Seek, SeekFrom, Write},
    sync::{Arc, Barrier},
    thread,
};

use litchi_core::{FileSource, FileVersionPolicy, ReadAt};
use tempfile::tempdir;

#[test]
fn reads_exact_ranges_and_reports_short_eof() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    fs::write(&path, b"abcdef").expect("write source");
    let source = FileSource::open(&path).expect("open regular file");

    assert_eq!(source.len().expect("source length"), 6);
    let mut exact = [0; 3];
    source.read_exact_at(2, &mut exact).expect("exact range");
    assert_eq!(&exact, b"cde");

    let mut short = [0; 4];
    assert_eq!(source.read_at(4, &mut short).expect("short read"), 2);
    assert_eq!(&short[..2], b"ef");
    let error = source
        .read_exact_at(4, &mut short)
        .expect_err("exact read crosses EOF");
    assert_eq!(error.kind(), io::ErrorKind::UnexpectedEof);

    assert_eq!(source.read_at(u64::MAX, &mut []).expect("empty read"), 0);
    let error = source
        .read_at(u64::MAX, &mut [0; 2])
        .expect_err("range overflows u64");
    assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
}

#[test]
fn clones_share_identity_and_support_concurrent_cursor_free_reads() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    let bytes: Vec<u8> = (0_u16..4096).map(|value| value as u8).collect();
    fs::write(&path, &bytes).expect("write source");
    let source = FileSource::open(&path).expect("open regular file");
    let expected_version = source.version().expect("initial version");
    let barrier = Arc::new(Barrier::new(9));

    let workers: Vec<_> = (0..8)
        .map(|worker| {
            let clone = source.clone();
            let barrier = Arc::clone(&barrier);
            thread::spawn(move || {
                barrier.wait();
                for round in 0..256 {
                    let offset = (worker * 257 + round * 13) % (4096 - 31);
                    let mut output = [0; 31];
                    clone
                        .read_exact_at(offset as u64, &mut output)
                        .expect("concurrent positional read");
                    for (index, byte) in output.into_iter().enumerate() {
                        assert_eq!(byte, (offset + index) as u8);
                    }
                }
                clone.version().expect("clone version")
            })
        })
        .collect();

    barrier.wait();
    for worker in workers {
        assert_eq!(worker.join().expect("worker completed"), expected_version);
    }
    assert_eq!(source.version().expect("stable version"), expected_version);
}

#[test]
fn truncation_changes_length_and_metadata_version() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    fs::write(&path, b"abcdefgh").expect("write source");
    let source = FileSource::open(&path).expect("open regular file");
    let before = source.version().expect("initial version");

    OpenOptions::new()
        .write(true)
        .open(&path)
        .expect("open writer")
        .set_len(3)
        .expect("truncate source");

    assert_eq!(source.len().expect("truncated length"), 3);
    let after = source.version().expect("changed version");
    assert_eq!(after.id(), before.id());
    assert!(after.revision() > before.revision());
    let error = source
        .read_exact_at(0, &mut [0; 4])
        .expect_err("truncated range crosses EOF");
    assert_eq!(error.kind(), io::ErrorKind::UnexpectedEof);
}

#[test]
fn replacing_path_does_not_retarget_an_open_source() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    let replacement = directory.path().join("replacement.bin");
    fs::write(&path, b"old bytes").expect("write old source");
    fs::write(&replacement, b"new bytes").expect("write replacement");
    let old_source = FileSource::open(&path).expect("open old source");

    fs::remove_file(&path).expect("unlink old path");
    fs::rename(&replacement, &path).expect("install replacement");
    let new_source = FileSource::open(&path).expect("open replacement source");
    let mut old = [0; 9];
    let mut new = [0; 9];
    old_source
        .read_exact_at(0, &mut old)
        .expect("read pinned old handle");
    new_source
        .read_exact_at(0, &mut new)
        .expect("read replacement handle");

    assert_eq!(&old, b"old bytes");
    assert_eq!(&new, b"new bytes");
    assert_ne!(
        old_source.version().expect("old source version").id(),
        new_source.version().expect("new source version").id()
    );
}

#[test]
fn accepts_owned_file_and_rejects_directory_paths() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    fs::write(&path, b"bytes").expect("write source");
    let source =
        FileSource::from_file(File::open(&path).expect("open file")).expect("accept regular file");
    let mut output = [0; 5];
    source.read_exact_at(0, &mut output).expect("read file");
    assert_eq!(&output, b"bytes");

    let error = FileSource::open(directory.path()).expect_err("reject directory");
    assert!(matches!(
        error.kind(),
        io::ErrorKind::InvalidInput | io::ErrorKind::PermissionDenied
    ));

    #[cfg(unix)]
    assert_eq!(source.version_policy(), FileVersionPolicy::UnixMetadata);
    #[cfg(windows)]
    assert_eq!(
        source.version_policy(),
        FileVersionPolicy::WindowsTimestampMetadata
    );
}

#[test]
fn positional_reads_do_not_move_a_shared_file_cursor() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    fs::write(&path, b"abcdef").expect("write source");
    let file = File::open(&path).expect("open file");
    let mut cursor_observer = file.try_clone().expect("clone file handle");
    cursor_observer
        .seek(SeekFrom::Start(3))
        .expect("position shared cursor");
    let source = FileSource::from_file(file).expect("construct positional source");

    let mut output = [0; 2];
    source.read_exact_at(0, &mut output).expect("read at zero");

    assert_eq!(&output, b"ab");
    assert_eq!(
        cursor_observer.stream_position().expect("observe cursor"),
        3
    );
}

#[cfg(unix)]
#[test]
fn open_explicitly_follows_symlinks_to_regular_files() {
    use std::os::unix::fs::symlink;

    let directory = tempdir().expect("temporary directory");
    let target = directory.path().join("target.bin");
    let link = directory.path().join("link.bin");
    fs::write(&target, b"target").expect("write target");
    symlink(&target, &link).expect("create symlink");

    let source = FileSource::open(&link).expect("follow regular-file symlink");
    let mut output = [0; 6];
    source.read_exact_at(0, &mut output).expect("read target");
    assert_eq!(&output, b"target");
    assert_eq!(source.version_policy(), FileVersionPolicy::UnixMetadata);
}

#[test]
fn writing_through_an_independent_handle_changes_the_version() {
    let directory = tempdir().expect("temporary directory");
    let path = directory.path().join("source.bin");
    fs::write(&path, b"old").expect("write source");
    let source = FileSource::open(&path).expect("open source");
    let before = source.version().expect("initial version");

    let mut writer = OpenOptions::new()
        .append(true)
        .open(&path)
        .expect("open writer");
    writer.write_all(b" bytes").expect("append bytes");
    writer.sync_all().expect("synchronize writer");
    drop(writer);

    let after = source.version().expect("changed version");
    assert_eq!(after.id(), before.id());
    assert_eq!(after.revision(), before.revision() + 1);
    assert_eq!(source.version().expect("stable changed version"), after);
}
