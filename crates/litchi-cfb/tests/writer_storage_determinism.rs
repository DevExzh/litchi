#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]
//! `OleWriter::write_to` must serialize explicitly created storages in an order
//! derived from the document, never from the process's hash seed.
//!
//! Change 0617 measured the defect these tests guard: building the same
//! eight-storage document in twelve separate processes produced twelve distinct
//! files of identical length, because `OleWriter` held its storage paths in a
//! `HashSet` whose iteration order is seeded per process
//! (`docs/performance/results/change-0617/determinism.txt`). ADR 0006 requires
//! serialization to be deterministic unless a `Clock`, actor identity, or
//! cryptographic RNG is explicitly supplied, and the writer is supplied none.
//!
//! Each `OleWriter::new()` builds its own `RandomState`, which Rust seeds from a
//! per-thread counter that advances on every construction, so repeating a build
//! in one process varies the hash order exactly as separate processes do. The
//! digests below are FNV-1a, the digest 0617's retained evidence uses, under the
//! standard 64-bit prime.

use litchi_cfb::writer::OleWriter;
use litchi_cfb::{OleError, OleFile};
use std::collections::BTreeSet;
use std::io::Cursor;

/// Repetitions per determinism assertion. Twelve processes were enough for
/// change 0617 to see twelve distinct digests at eight storages; sixty-four
/// in-process builds make an accidental agreement vanishingly unlikely.
const REPEATS: usize = 64;

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}

/// Builds change 0617's probe document: one explicitly created storage per
/// index, each holding one `Payload` stream. The storages are declared in
/// `order`; the streams are always declared in ascending index order, because
/// stream insertion order is a deliberate part of the writer's model (a DOC's
/// `WordDocument` must be added first to reach sector 0) and is not what these
/// tests vary.
fn sibling_storages(order: &[usize]) -> Result<Vec<u8>, OleError> {
    let mut writer = OleWriter::new();
    for index in order {
        writer.create_storage(&[format!("Storage{index:02}").as_str()])?;
    }
    let mut ascending: Vec<usize> = order.to_vec();
    ascending.sort_unstable();
    for index in ascending {
        let name = format!("Storage{index:02}");
        writer.create_stream(
            &[name.as_str(), "Payload"],
            format!("payload-{index}").as_bytes(),
        )?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

/// Builds two two-level subtrees whose storages are declared child-before-parent
/// and out of sibling order, so the canonical order has to place every ancestor
/// ahead of its descendants on its own.
fn nested_storages(order: &[&[&str]]) -> Result<Vec<u8>, OleError> {
    let mut writer = OleWriter::new();
    for path in order {
        writer.create_storage(path)?;
    }
    for path in order {
        let mut stream_path = path.to_vec();
        stream_path.push("Payload");
        writer.create_stream(&stream_path, path.join("/").as_bytes())?;
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn digest(bytes: &[u8]) -> (usize, u64) {
    (bytes.len(), fnv1a(bytes))
}

#[test]
fn eight_sibling_storages_serialize_to_one_digest_across_hasher_seeds() {
    let order: Vec<usize> = (0..8).collect();
    let mut digests = BTreeSet::new();
    for _ in 0..REPEATS {
        digests.insert(digest(&sibling_storages(&order).unwrap()));
    }
    assert_eq!(
        digests.len(),
        1,
        "eight explicitly created storages serialized to {} distinct outputs in {REPEATS} builds: {digests:?}",
        digests.len()
    );
}

#[test]
fn nested_storages_serialize_to_one_digest_across_hasher_seeds() {
    let order: [&[&str]; 6] = [
        &["Beta", "Inner"],
        &["Beta"],
        &["Alpha", "Second"],
        &["Alpha", "First"],
        &["Alpha"],
        &["Gamma"],
    ];
    let mut digests = BTreeSet::new();
    for _ in 0..REPEATS {
        digests.insert(digest(&nested_storages(&order).unwrap()));
    }
    assert_eq!(
        digests.len(),
        1,
        "six nested storages serialized to {} distinct outputs in {REPEATS} builds: {digests:?}",
        digests.len()
    );
}

#[test]
fn storage_declaration_order_does_not_change_the_bytes() {
    let ascending: Vec<usize> = (0..8).collect();
    let descending: Vec<usize> = (0..8).rev().collect();
    let interleaved: Vec<usize> = (0..8).map(|index| (index * 3) % 8).collect();

    let reference = sibling_storages(&ascending).unwrap();
    assert_eq!(
        digest(&sibling_storages(&descending).unwrap()),
        digest(&reference),
        "reversing the declaration order changed the serialized bytes"
    );
    assert_eq!(
        digest(&sibling_storages(&interleaved).unwrap()),
        digest(&reference),
        "permuting the declaration order changed the serialized bytes"
    );
}

#[test]
fn single_storage_output_is_byte_identical_to_the_pre_fix_writer() {
    // Change 0617 captured this document from the unmodified writer, where one
    // storage was already deterministic: 2,560 bytes in twelve of twelve runs.
    // Its probe printed `83490167a5417d81`, having multiplied by sixteen times
    // the FNV-1a prime; the same bytes digest to the constant below under the
    // standard prime, and change 0625 reproduced both. A canonical storage
    // order must leave these bytes exactly where they were.
    assert_eq!(
        digest(&sibling_storages(&[0]).unwrap()),
        (2560, 0xb795_f367_a541_7d81),
        "the one-storage output no longer matches change 0617's pre-fix digest"
    );
}

#[test]
fn the_deterministic_output_is_a_conforming_compound_file() {
    let order: Vec<usize> = (0..8).rev().collect();
    let bytes = sibling_storages(&order).unwrap();
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();

    let mut streams = file.list_streams();
    streams.sort();
    let expected: Vec<Vec<String>> = (0..8)
        .map(|index| vec![format!("Storage{index:02}"), "Payload".to_string()])
        .collect();
    assert_eq!(streams, expected);

    for index in 0..8 {
        let name = format!("Storage{index:02}");
        assert!(file.directory_exists(&[name.as_str()]));
        assert_eq!(
            file.open_stream(&[name.as_str(), "Payload"]).unwrap(),
            format!("payload-{index}").as_bytes()
        );
    }
}

#[test]
fn the_deterministic_nested_output_is_a_conforming_compound_file() {
    let order: [&[&str]; 6] = [
        &["Beta", "Inner"],
        &["Beta"],
        &["Alpha", "Second"],
        &["Alpha", "First"],
        &["Alpha"],
        &["Gamma"],
    ];
    let bytes = nested_storages(&order).unwrap();
    let mut file = OleFile::open(Cursor::new(bytes)).unwrap();

    for path in order {
        assert!(
            file.directory_exists(path),
            "storage {path:?} is missing from the serialized directory"
        );
        let mut stream_path = path.to_vec();
        stream_path.push("Payload");
        assert_eq!(
            file.open_stream(&stream_path).unwrap(),
            path.join("/").as_bytes()
        );
    }
}
