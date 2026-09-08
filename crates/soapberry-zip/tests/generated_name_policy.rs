//! Cross-route publication tests for checked generated member sequences.

use soapberry_zip::generated_names::{
    GeneratedNamePlan, GeneratedNamePlanBuilder, GeneratedNamePlanLimits,
};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveLimits, StreamingArchiveWriter};
use soapberry_zip::{CompressionMethod, DirectorySpoolLimits, ZipOperationAccounting};
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};

fn builder() -> GeneratedNamePlanBuilder {
    GeneratedNamePlanBuilder::new(GeneratedNamePlanLimits {
        max_patterns: 64,
        max_pattern_bytes: 4096,
        max_entries: 64,
    })
    .unwrap()
}

fn literals(names: &[&str]) -> GeneratedNamePlan {
    let mut builder = builder();
    for name in names {
        builder.push_literal(name).unwrap();
    }
    builder.finish().unwrap()
}

fn planned<W: Write>(output: W, plan: GeneratedNamePlan) -> StreamingArchiveWriter<W> {
    StreamingArchiveWriter::with_writer_and_limits_and_spool_and_name_plan(
        output,
        StreamingArchiveLimits::default(),
        Cursor::new(Vec::new()),
        DirectorySpoolLimits::new(1 << 20, 17),
        plan,
    )
    .unwrap()
}

fn mixed(mut archive: StreamingArchiveWriter<Vec<u8>>) -> Vec<u8> {
    archive.write_stored("metadata.bin", b"metadata").unwrap();
    let mut entry = archive
        .start_entry("items/part1.xml", CompressionMethod::Deflate)
        .unwrap();
    entry.write_all(b"first payload").unwrap();
    archive = entry.finish().unwrap();
    archive
        .write_stored_stream("items/_rels/part1.xml.rels", &b"first relation"[..])
        .unwrap();
    archive
        .write_deflated_sized("items/part2.xml", b"second payload")
        .unwrap();
    archive
        .write_deflated_with_accounting(
            "items/_rels/part2.xml.rels",
            b"second relation",
            &mut ZipOperationAccounting::default(),
        )
        .unwrap();
    archive
        .write_stream("tail.bin", &b"tail"[..], CompressionMethod::Store)
        .unwrap();
    archive.finish().unwrap()
}

#[test]
fn mixed_owned_sized_and_reader_routes_preserve_bytes_and_payloads() {
    let mut plan = builder();
    plan.push_literal("metadata.bin").unwrap();
    plan.push_indexed(
        1,
        2,
        &[("items/part", ".xml"), ("items/_rels/part", ".xml.rels")],
    )
    .unwrap();
    plan.push_literal("tail.bin").unwrap();
    let planned = mixed(planned(Vec::new(), plan.finish().unwrap()));
    assert_eq!(
        planned,
        mixed(StreamingArchiveWriter::with_writer(Vec::new()))
    );
    let archive = ArchiveReader::new(&planned).unwrap();
    assert_eq!(archive.len(), 6);
    for (name, expected) in [
        ("metadata.bin", &b"metadata"[..]),
        ("items/part1.xml", &b"first payload"[..]),
        ("items/_rels/part1.xml.rels", &b"first relation"[..]),
        ("items/part2.xml", &b"second payload"[..]),
        ("items/_rels/part2.xml.rels", &b"second relation"[..]),
        ("tail.bin", &b"tail"[..]),
    ] {
        assert_eq!(archive.read(name).unwrap(), expected);
    }
}

struct MustNotRead;

impl Read for MustNotRead {
    fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
        panic!("name rejection must precede payload reads");
    }
}

struct MustNotAccessSpool;

impl Read for MustNotAccessSpool {
    fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
        panic!("plan limits must be checked before scratch reads");
    }
}

impl Write for MustNotAccessSpool {
    fn write(&mut self, _: &[u8]) -> io::Result<usize> {
        panic!("plan limits must be checked before scratch writes");
    }

    fn flush(&mut self) -> io::Result<()> {
        panic!("plan limits must be checked before scratch flushes");
    }
}

impl Seek for MustNotAccessSpool {
    fn seek(&mut self, _: SeekFrom) -> io::Result<u64> {
        panic!("plan limits must be checked before scratch initialization");
    }
}

#[test]
fn every_borrowed_route_refuses_wrong_names_without_consuming_the_plan() {
    for route in 0..32 {
        for bad_name in [
            "different.bin",
            "/expected1.bin",
            "expected1.bin/",
            "EXPECTED1.BIN",
            "EXPECTED1.bin",
            "expected1.BIN",
            "expected1.bin/child",
        ] {
            let plan = if route < 16 {
                literals(&["expected1.bin"])
            } else {
                let mut plan = builder();
                plan.push_indexed(1, 1, &[("expected", ".bin")]).unwrap();
                plan.finish().unwrap()
            };
            let mut archive = planned(Vec::new(), plan);
            let mut accounting = ZipOperationAccounting::default();
            let result = match route % 16 {
                0 => archive.write_stored(bad_name, b"x"),
                1 => archive.write_stored_with_accounting(bad_name, b"x", &mut accounting),
                2 => archive.write_deflated(bad_name, b"x"),
                3 => archive.write_deflated_with_accounting(bad_name, b"x", &mut accounting),
                4 => archive.write_deflated_sized(bad_name, b"x"),
                5 => archive.write_deflated_sized_with_accounting(bad_name, b"x", &mut accounting),
                6 => archive.write_stored_stream(bad_name, MustNotRead),
                7 => archive.write_stored_stream_with_accounting(
                    bad_name,
                    MustNotRead,
                    &mut accounting,
                ),
                8 => archive.write_deflated_stream(bad_name, MustNotRead),
                9 => archive.write_deflated_stream_with_accounting(
                    bad_name,
                    MustNotRead,
                    &mut accounting,
                ),
                10 => archive.write_stored_reader(bad_name, MustNotRead),
                11 => archive.write_stored_reader_with_accounting(
                    bad_name,
                    MustNotRead,
                    &mut accounting,
                ),
                12 => archive.write_deflated_reader(bad_name, MustNotRead),
                13 => archive.write_deflated_reader_with_accounting(
                    bad_name,
                    MustNotRead,
                    &mut accounting,
                ),
                14 => archive.write_stream(bad_name, MustNotRead, CompressionMethod::Store),
                15 => archive.write_stream_with_accounting(
                    bad_name,
                    MustNotRead,
                    CompressionMethod::Deflate,
                    &mut accounting,
                ),
                _ => unreachable!(),
            };
            assert!(result.is_err(), "route {route}: {bad_name}");
            assert_eq!(archive.output_bytes(), 0);
            archive.write_stored("expected1.bin", b"accepted").unwrap();
            let output = archive.finish().unwrap();
            assert_eq!(
                ArchiveReader::new(&output)
                    .unwrap()
                    .read("expected1.bin")
                    .unwrap(),
                b"accepted"
            );
        }
    }
}

#[test]
fn owned_route_refuses_skipped_or_repeated_members_before_header_output() {
    let mut output = Vec::new();
    let archive = planned(&mut output, literals(&["first.bin", "second.bin"]));
    assert!(
        archive
            .start_entry("second.bin", CompressionMethod::Store)
            .is_err()
    );
    assert!(output.is_empty());

    let mut archive = planned(&mut output, literals(&["first.bin", "second.bin"]));
    archive.write_stored("first.bin", b"first").unwrap();
    let before = archive.output_bytes();
    assert!(
        archive
            .start_entry("first.bin", CompressionMethod::Deflate)
            .is_err()
    );
    assert_eq!(output.len() as u64, before);
}

#[test]
fn missing_members_refuse_finalization_with_exact_progress() {
    let mut output = Vec::new();
    let mut archive = planned(&mut output, literals(&["first.bin", "second.bin"]));
    archive.write_stored("first.bin", b"first").unwrap();
    let before = archive.output_bytes();
    let failure = archive.finish_with_progress().unwrap_err();
    assert_eq!(failure.progress().output_bytes(), before);
    assert_eq!(output.len() as u64, before);
    assert!(ArchiveReader::new(&output).is_err());
}

#[test]
fn a_completed_plan_rejects_extra_members_but_can_finish() {
    let mut archive = planned(Vec::new(), literals(&["only.bin"]));
    archive.write_stored("only.bin", b"one").unwrap();
    let before = archive.output_bytes();
    assert!(archive.write_stored("extra.bin", b"extra").is_err());
    assert_eq!(archive.output_bytes(), before);
    assert_eq!(
        ArchiveReader::new(&archive.finish().unwrap())
            .unwrap()
            .len(),
        1
    );
}

#[test]
fn generated_plan_cannot_raise_streaming_entry_or_name_limits() {
    for limits in [
        StreamingArchiveLimits {
            max_entries: 0,
            ..StreamingArchiveLimits::default()
        },
        StreamingArchiveLimits {
            max_member_name_bytes: 2,
            ..StreamingArchiveLimits::default()
        },
    ] {
        let mut output = Vec::new();
        let result = StreamingArchiveWriter::with_writer_and_limits_and_spool_and_name_plan(
            &mut output,
            limits,
            MustNotAccessSpool,
            DirectorySpoolLimits::new(1 << 20, 17),
            literals(&["only.bin"]),
        );
        assert!(result.is_err());
        drop(result);
        assert!(output.is_empty());
    }
}

#[test]
fn an_empty_plan_can_finalize_an_empty_archive() {
    let output = planned(Vec::new(), builder().finish().unwrap())
        .finish()
        .unwrap();
    assert_eq!(ArchiveReader::new(&output).unwrap().len(), 0);
}

#[test]
fn ordinary_name_validation_still_rejects_duplicates() {
    let mut archive = StreamingArchiveWriter::with_writer(Vec::new());
    archive.write_stored("a/b.bin", b"first").unwrap();
    let before = archive.output_bytes();
    assert!(archive.write_stored("a/b.bin", b"duplicate").is_err());
    assert_eq!(archive.output_bytes(), before);
    archive.write_stored("different.bin", b"second").unwrap();
    assert_eq!(
        ArchiveReader::new(&archive.finish().unwrap())
            .unwrap()
            .len(),
        2
    );
}
