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

fn directory_image_offset(bytes: &[u8]) -> usize {
    let sector_size = 1usize << u16::from_le_bytes(bytes[0x1E..0x20].try_into().unwrap());
    let directory_sector = u32::from_le_bytes(bytes[0x30..0x34].try_into().unwrap()) as usize;
    (directory_sector + 1) * sector_size
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

#[test]
fn shape_fallback_preserves_adopted_clsids_until_explicitly_changed() {
    let mut source_writer = OleWriter::new();
    source_writer.set_root_clsid([0x11; 16]);
    source_writer.create_storage(&["Object"]).unwrap();
    source_writer
        .set_storage_clsid(&["Object"], [0x22; 16])
        .unwrap();
    source_writer
        .create_stream(&["Object", "Payload"], &[0x33; 9000])
        .unwrap();
    let mut source_output = Cursor::new(Vec::new());
    source_writer.write_to(&mut source_output).unwrap();
    let source = source_output.into_inner();

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer.create_storage(&["Object"]).unwrap();
    writer.create_storage(&["Added"]).unwrap();
    writer
        .create_stream(&["Object", "Payload"], &[0x44; 9000])
        .unwrap();
    writer
        .create_stream(&["Added", "Payload"], &[0x55; 32])
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let output = output.into_inner();
    assert_eq!(
        writer.last_sector_layout().unwrap().fallback(),
        Some(litchi_cfb::SectorLayoutFallback::DirectoryShapeChanged)
    );

    let ole = OleFile::open(Cursor::new(output)).unwrap();
    assert!(ole.root_entry().unwrap().clsid.contains("11111111"));
    let object = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Object")
        .unwrap();
    assert!(object.clsid.contains("22222222"));
}

#[test]
fn explicit_zero_clsids_clear_source_directory_fields() {
    let mut source_writer = OleWriter::new();
    source_writer.set_root_clsid([0x11; 16]);
    source_writer.create_storage(&["Object"]).unwrap();
    source_writer
        .set_storage_clsid(&["Object"], [0x22; 16])
        .unwrap();
    source_writer
        .create_stream(&["Object", "Payload"], &[0x33; 9000])
        .unwrap();
    let mut source_output = Cursor::new(Vec::new());
    source_writer.write_to(&mut source_output).unwrap();
    let source = source_output.into_inner();

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&source).unwrap());
    writer.set_root_clsid([0; 16]);
    writer.create_storage(&["Object"]).unwrap();
    writer.set_storage_clsid(&["Object"], [0; 16]).unwrap();
    writer
        .create_stream(&["Object", "Payload"], &[0x33; 9000])
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let output = output.into_inner();
    assert!(writer.last_sector_layout().unwrap().reused_source_layout());

    let ole = OleFile::open(Cursor::new(output.clone())).unwrap();
    assert!(ole.root_entry().unwrap().clsid.is_empty());
    let entries = ole.list_directory_entries(&[]).unwrap();
    let object = entries.iter().find(|entry| entry.name == "Object").unwrap();
    assert!(object.clsid.is_empty());
    let directory = directory_image_offset(&output);
    assert_eq!(&output[directory + 0x50..directory + 0x60], &[0; 16]);
    let object_offset = directory + object.sid as usize * 128;
    assert_eq!(
        &output[object_offset + 0x50..object_offset + 0x60],
        &[0; 16]
    );
}

#[test]
fn reused_v3_zero_length_stream_masks_high_size_word() {
    let source = build(&[(&["Empty"], Vec::new())]);
    let ole = OleFile::open(Cursor::new(source.clone())).unwrap();
    let sid = ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Empty")
        .unwrap()
        .sid;
    let mut mutated = source;
    let directory = directory_image_offset(&mutated);
    let size_offset = directory + sid as usize * 128 + 120;
    mutated[size_offset..size_offset + 8].copy_from_slice(&0xDEAD_BEEF_0000_0000u64.to_le_bytes());

    let mut writer = OleWriter::new();
    assert!(writer.adopt_source_layout(&mutated).unwrap());
    writer.create_stream(&["Empty"], &[]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let output = output.into_inner();
    assert!(writer.last_sector_layout().unwrap().reused_source_layout());
    let directory = directory_image_offset(&output);
    let size_offset = directory + sid as usize * 128 + 120;
    assert_eq!(&output[size_offset..size_offset + 8], &[0; 8]);
    OleFile::open(Cursor::new(output)).unwrap();
}

#[test]
fn shrinking_a_mini_stream_clears_former_payload_from_root_padding() {
    let original = vec![1u8; 100];
    let source = build(&[(&["Small"], original.clone())]);
    let mut expanded = original.clone();
    expanded.extend_from_slice(&[9u8; 70]);
    let render = |source: &[u8], payload: &[u8]| {
        let mut writer = OleWriter::new();
        assert!(writer.adopt_source_layout(source).unwrap());
        writer.create_stream(&["Small"], payload).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        assert!(writer.last_sector_layout().unwrap().reused_source_layout());
        output.into_inner()
    };
    let grown = render(&source, &expanded);
    assert_eq!(
        grown.len(),
        source.len(),
        "growth fits the root's last sector"
    );
    let restored = render(&grown, &original);
    assert_eq!(read_all(&restored), read_all(&source));
    assert_eq!(restored, source, "released root padding must be zero again");
}
