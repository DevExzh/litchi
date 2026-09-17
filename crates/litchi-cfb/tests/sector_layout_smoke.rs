//! Smoke checks for the reuse sector-layout policy.

use litchi_cfb::{OleFile, OleWriter, SectorLayoutPolicy};
use std::io::Cursor;

fn build(streams: &[(&[&str], Vec<u8>)]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        writer.create_stream(path, data).unwrap();
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    out.into_inner()
}

fn read_all(bytes: &[u8]) -> Vec<(Vec<String>, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let mut paths = ole.list_streams();
    paths.sort();
    paths
        .into_iter()
        .map(|path| {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&refs).unwrap();
            (path, data)
        })
        .collect()
}

#[test]
fn reuse_keeps_sectors_for_a_same_length_edit() {
    let source = build(&[
        (&["WordDocument"], vec![1u8; 9000]),
        (&["1Table"], vec![2u8; 5000]),
        (&["Small"], vec![3u8; 100]),
    ]);

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer
        .create_stream(&["WordDocument"], &vec![9u8; 9000])
        .unwrap();
    writer.create_stream(&["1Table"], &vec![2u8; 5000]).unwrap();
    writer.create_stream(&["Small"], &[3u8; 100]).unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let reused = out.into_inner();
    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert_eq!(report.appended_sectors(), 0, "{report:?}");
    assert_eq!(reused.len(), source.len());

    let streams = read_all(&reused);
    assert_eq!(streams.len(), 3);
    assert_eq!(streams[2].1, vec![9u8; 9000]);
    assert_eq!(streams[0].1, vec![2u8; 5000]);
    assert_eq!(streams[1].1, vec![3u8; 100]);
}

#[test]
fn reuse_appends_when_a_stream_outgrows_its_allocation() {
    let source = build(&[
        (&["WordDocument"], vec![1u8; 9000]),
        (&["1Table"], vec![2u8; 5000]),
        (&["Small"], vec![3u8; 100]),
    ]);

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer
        .create_stream(&["WordDocument"], &vec![9u8; 40000])
        .unwrap();
    writer.create_stream(&["1Table"], &vec![2u8; 5000]).unwrap();
    writer.create_stream(&["Small"], &[3u8; 100]).unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let reused = out.into_inner();
    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert!(report.appended_sectors() > 0, "{report:?}");

    let streams = read_all(&reused);
    assert_eq!(streams[2].1, vec![9u8; 40000]);
    assert_eq!(streams[0].1, vec![2u8; 5000]);
    assert_eq!(streams[1].1, vec![3u8; 100]);
}

#[test]
fn reuse_handles_mini_to_regular_cutoff_migration() {
    let source = build(&[(&["Payload"], vec![0x11u8; 4095])]);

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer
        .create_stream(&["Payload"], &vec![0x22u8; 5000])
        .unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let output = out.into_inner();
    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert!(report.reclaimed_sectors() > 0, "{report:?}");
    assert_eq!(
        read_all(&output),
        vec![(vec!["Payload".to_string()], vec![0x22; 5000])]
    );
}

#[test]
fn reuse_handles_regular_to_mini_cutoff_migration() {
    let source = build(&[(&["Payload"], vec![0x11u8; 5000])]);

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer
        .create_stream(&["Payload"], &vec![0x22u8; 4095])
        .unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let output = out.into_inner();
    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert!(report.reclaimed_sectors() > 0, "{report:?}");
    assert_eq!(
        read_all(&output),
        vec![(vec!["Payload".to_string()], vec![0x22; 4095])]
    );
}

#[test]
fn rewrite_policy_declines_and_matches_today() {
    let source = build(&[
        (&["WordDocument"], vec![1u8; 9000]),
        (&["1Table"], vec![2u8; 5000]),
    ]);

    let mut plain = OleWriter::new();
    plain
        .create_stream(&["WordDocument"], &vec![9u8; 9000])
        .unwrap();
    plain.create_stream(&["1Table"], &vec![2u8; 5000]).unwrap();
    let mut a = Cursor::new(Vec::new());
    plain.write_to(&mut a).unwrap();

    let mut opted = OleWriter::new();
    opted.set_sector_layout_policy(SectorLayoutPolicy::Rewrite);
    assert!(opted.adopt_source_layout(&source).unwrap());
    opted
        .create_stream(&["WordDocument"], &vec![9u8; 9000])
        .unwrap();
    opted.create_stream(&["1Table"], &vec![2u8; 5000]).unwrap();
    let mut b = Cursor::new(Vec::new());
    opted.write_to(&mut b).unwrap();

    assert_eq!(a.into_inner(), b.into_inner());
    assert!(!opted.last_sector_layout().unwrap().reused_source_layout());
}

#[test]
fn shrinking_a_stream_reclaims_its_sectors() {
    let source = build(&[
        (&["WordDocument"], vec![1u8; 40000]),
        (&["1Table"], vec![2u8; 5000]),
    ]);

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer
        .create_stream(&["WordDocument"], &vec![9u8; 6000])
        .unwrap();
    writer
        .create_stream(&["1Table"], &vec![2u8; 60000])
        .unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let reused = out.into_inner();
    let report = writer.last_sector_layout().unwrap();
    assert!(report.reused_source_layout(), "{report:?}");
    assert!(report.reclaimed_sectors() > 0, "{report:?}");

    let streams = read_all(&reused);
    assert_eq!(streams[1].1, vec![9u8; 6000]);
    assert_eq!(streams[0].1, vec![2u8; 60000]);
}

#[test]
fn a_changed_stream_set_declines() {
    let source = build(&[(&["A"], vec![1u8; 9000])]);
    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer.create_stream(&["A"], &vec![1u8; 9000]).unwrap();
    writer.create_stream(&["B"], &vec![2u8; 9000]).unwrap();
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).unwrap();
    let report = writer.last_sector_layout().unwrap();
    assert!(!report.reused_source_layout());
    assert_eq!(
        report.fallback(),
        Some(litchi_cfb::SectorLayoutFallback::DirectoryShapeChanged)
    );
    assert_eq!(read_all(&out.into_inner()).len(), 2);
}
