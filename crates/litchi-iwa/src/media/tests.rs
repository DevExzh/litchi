//! Focused media discovery and transactional-edit regressions.

use std::fs;
use std::sync::Arc;

use sha1::{Digest, Sha1};

use crate::archive::{Archive, ArchiveObject, RawMessage};
use crate::package::{IWorkPackage, PackageLimits};
use litchi_iwa_common::varint::encode_varint;

use super::*;

const PNG: &[u8] = b"\x89PNG\r\n\x1a\nreplacement";

fn asset_id(raw: u64) -> MediaAssetId {
    MediaAssetId::new(raw).expect("test media asset identifiers are non-zero")
}

fn append_varint_field(output: &mut Vec<u8>, number: u64, value: u64) {
    output.extend(encode_varint(number << 3));
    output.extend(encode_varint(value));
}

fn append_bytes_field(output: &mut Vec<u8>, number: u64, value: &[u8]) {
    output.extend(encode_varint((number << 3) | 2));
    output.extend(encode_varint(value.len() as u64));
    output.extend_from_slice(value);
}

fn synthetic_metadata(data_bytes: &[u8]) -> Vec<u8> {
    let mut data_info = Vec::new();
    append_varint_field(&mut data_info, 1, 7);
    append_bytes_field(&mut data_info, 2, &[0x11; 20]);
    append_bytes_field(&mut data_info, 3, b"image.png");
    append_bytes_field(&mut data_info, 4, b"image-7.png");
    // DataAttributes is an empty generated message, but Apple writes
    // extension fields inside it. This payload must remain byte-exact.
    append_bytes_field(&mut data_info, 10, &[0x08, 0x96, 0x01, 0x12, 0x01, 0xff]);
    append_varint_field(&mut data_info, 18, data_bytes.len() as u64);

    let mut metadata = Vec::new();
    append_varint_field(&mut metadata, 1, 100);
    append_bytes_field(&mut metadata, 4, &data_info);
    append_bytes_field(&mut metadata, 100, b"outer-unknown");
    metadata
}

fn synthetic_package() -> IWorkPackage {
    let original = b"\x89PNG\r\n\x1a\noriginal";
    let mut metadata_object = ArchiveObject::new(
        2,
        vec![RawMessage {
            type_: PACKAGE_METADATA_MESSAGE_TYPE,
            data: synthetic_metadata(original),
        }],
    )
    .unwrap();
    metadata_object.archive_info.message_infos[0].data_references = Vec::new();
    let metadata_archive = Archive {
        objects: vec![metadata_object],
    };

    let mut document_object = ArchiveObject::new(
        50,
        vec![RawMessage {
            type_: 999,
            data: vec![1],
        }],
    )
    .unwrap();
    document_object.archive_info.message_infos[0].data_references = vec![7];
    let document_archive = Archive {
        objects: vec![document_object],
    };

    let mut package = IWorkPackage::new();
    package
        .replace_archive(PACKAGE_METADATA_ENTRY, &metadata_archive)
        .unwrap();
    package
        .replace_archive("Index/Document.iwa", &document_archive)
        .unwrap();
    package
        .insert_entry("Data/image-7.png", original.to_vec())
        .unwrap();
    package
}

fn nested_field_bytes(metadata: &[u8], field_number: u32) -> Vec<u8> {
    let outer = parse_wire_fields(metadata).unwrap();
    let data_info = outer
        .iter()
        .find(|field| field.number == 4)
        .map(|field| field_payload(metadata, field).unwrap())
        .unwrap();
    let nested = parse_wire_fields(data_info).unwrap();
    let field = nested
        .iter()
        .find(|field| field.number == field_number)
        .unwrap();
    data_info[field.start..field.end].to_vec()
}

#[test]
fn manager_reads_single_file_and_memory_packages() {
    let package = synthetic_package();
    let bytes = package.to_bytes().unwrap();
    let memory = MediaManager::from_bytes(&bytes).unwrap();
    assert_eq!(memory.assets().len(), 1);
    assert!(memory.get("image-7.png").unwrap().is_image());
    assert_eq!(
        memory.extract("image-7.png").unwrap(),
        b"\x89PNG\r\n\x1a\noriginal"
    );

    let file = tempfile::NamedTempFile::new().unwrap();
    fs::write(file.path(), bytes).unwrap();
    let disk = MediaManager::new(file.path()).unwrap();
    assert_eq!(
        disk.extract("image-7.png").unwrap(),
        memory.extract("image-7.png").unwrap()
    );

    let mut streamed = Vec::new();
    memory
        .extract_to_writer("image-7.png", &mut streamed)
        .unwrap();
    assert_eq!(streamed, b"\x89PNG\r\n\x1a\noriginal");
}

#[test]
fn media_file_extraction_rejects_non_file_destinations() -> std::io::Result<()> {
    let directory = tempfile::tempdir()?;
    let manager = MediaManager::from_package(synthetic_package()).unwrap();

    let error = manager
        .extract_to_file("image-7.png", directory.path())
        .unwrap_err();
    assert!(error.to_string().contains("not a regular file"));
    Ok(())
}

#[cfg(unix)]
#[test]
fn media_file_extraction_rejects_symbolic_link_destinations() -> std::io::Result<()> {
    use std::os::unix::fs::symlink;

    let directory = tempfile::tempdir()?;
    let target = directory.path().join("target.png");
    let link = directory.path().join("image.png");
    fs::write(&target, b"sentinel")?;
    symlink(&target, &link)?;

    let manager = MediaManager::from_package(synthetic_package()).unwrap();
    let error = manager.extract_to_file("image-7.png", &link).unwrap_err();

    assert!(error.to_string().contains("symbolic link"));
    assert_eq!(fs::read(&target)?, b"sentinel");
    assert!(fs::symlink_metadata(&link)?.file_type().is_symlink());
    Ok(())
}

#[cfg(unix)]
#[test]
fn media_file_extraction_preserves_existing_permissions() -> std::io::Result<()> {
    use std::os::unix::fs::PermissionsExt;

    let directory = tempfile::tempdir()?;
    let destination = directory.path().join("image.png");
    fs::write(&destination, b"old")?;
    let mut permissions = fs::metadata(&destination)?.permissions();
    permissions.set_mode(0o640);
    fs::set_permissions(&destination, permissions)?;

    let manager = MediaManager::from_package(synthetic_package()).unwrap();
    manager
        .extract_to_file("image-7.png", &destination)
        .unwrap();

    let mode = fs::metadata(&destination)?.permissions().mode() & 0o777;
    assert_eq!(mode, 0o640);
    assert_eq!(fs::read(&destination)?, b"\x89PNG\r\n\x1a\noriginal");
    Ok(())
}

#[test]
fn directory_media_streaming_rechecks_growth() -> std::io::Result<()> {
    let bundle = tempfile::tempdir()?;
    let data = bundle.path().join("Data");
    fs::create_dir(&data)?;
    let asset_path = data.join("asset.bin");
    fs::write(&asset_path, b"small")?;

    let limits = MediaLimits::new(1, 5, 5).unwrap();
    let manager = MediaManager::new_with_limits(bundle.path(), limits).unwrap();
    fs::write(&asset_path, b"larger")?;

    let mut streamed = Vec::new();
    let error = manager
        .extract_to_writer("asset.bin", &mut streamed)
        .unwrap_err();
    assert!(error.to_string().contains("grew to"));
    assert!(streamed.is_empty());

    let destination = bundle.path().join("output.bin");
    fs::write(&destination, b"sentinel")?;
    assert!(manager.extract_to_file("asset.bin", &destination).is_err());
    assert_eq!(fs::read(destination)?, b"sentinel");
    Ok(())
}

#[test]
fn filtered_media_queries_are_deterministic_by_relative_path() {
    let mut package = synthetic_package();
    package
        .insert_entry("Data/z-image.png", b"z".to_vec())
        .unwrap();
    package
        .insert_entry("Data/a-image.png", b"a".to_vec())
        .unwrap();
    package
        .insert_entry("Data/clip.m4a", b"m4a".to_vec())
        .unwrap();

    let manager = MediaManager::from_package(package).unwrap();
    let all_paths: Vec<_> = manager
        .assets_in_order()
        .iter()
        .map(|asset| asset.path.to_string_lossy().into_owned())
        .collect();
    assert_eq!(
        all_paths,
        vec![
            "Data/a-image.png",
            "Data/clip.m4a",
            "Data/image-7.png",
            "Data/z-image.png"
        ]
    );

    let image_paths: Vec<_> = manager
        .images()
        .into_iter()
        .map(|asset| asset.path.to_string_lossy().into_owned())
        .collect();
    assert_eq!(
        image_paths,
        vec!["Data/a-image.png", "Data/image-7.png", "Data/z-image.png"]
    );

    let audio_paths: Vec<_> = manager
        .assets_by_type(MediaType::Audio)
        .into_iter()
        .map(|asset| asset.path.to_string_lossy().into_owned())
        .collect();
    assert_eq!(audio_paths, vec!["Data/clip.m4a"]);
}

fn assert_send_sync<T: Send + Sync>() {}

#[test]
fn media_manager_snapshots_are_shared_and_thread_safe() {
    assert_send_sync::<MediaManager>();

    let package = synthetic_package();
    let manager = MediaManager::from_package(package).unwrap();
    let snapshot = manager.snapshot();

    assert!(Arc::ptr_eq(&manager.state, &snapshot.state));
    assert_eq!(snapshot.stats(), manager.stats());
    assert_eq!(
        snapshot.extract("image-7.png").unwrap(),
        manager.extract("image-7.png").unwrap()
    );
}

#[test]
fn media_limits_bound_discovery_and_extraction() {
    let package = synthetic_package();
    let bytes = package.to_bytes().unwrap();
    let tight = MediaLimits::new(1, 1, 1).unwrap();
    assert!(MediaManager::from_bytes_with_limits(&bytes, tight).is_err());

    let limits = MediaLimits::new(1, PNG.len() as u64, PNG.len() as u64).unwrap();
    let manager = MediaManager::from_package_with_limits(package, limits).unwrap();
    assert_eq!(manager.limits().max_assets(), 1);
    assert_eq!(manager.limits().max_asset_bytes(), PNG.len() as u64);
    assert_eq!(manager.limits().max_total_bytes(), PNG.len() as u64);

    assert!(MediaLimits::new(0, 1, 1).is_err());
    assert!(MediaLimits::new(1, 0, 1).is_err());
    assert!(MediaLimits::new(1, 1, 0).is_err());
    assert!(MediaLimits::new(MediaLimits::HARD_MAX_ASSETS + 1, 1, 1).is_err());
    assert!(MediaLimits::new(1, MediaLimits::HARD_MAX_ASSET_BYTES + 1, 1).is_err());
    assert!(MediaLimits::new(1, 1, MediaLimits::HARD_MAX_TOTAL_BYTES + 1).is_err());
}

#[test]
fn package_limits_are_reused_for_media_discovery_and_later_file_reads() {
    let package = synthetic_package();
    let bytes = package.to_bytes().unwrap();
    let media_limits = MediaLimits::new(1, 1024, 1024).unwrap();
    let package_limits = PackageLimits::new(
        PackageLimits::MAX_ENTRIES,
        PackageLimits::MAX_ENTRY_BYTES,
        PackageLimits::MAX_TOTAL_BYTES,
    )
    .unwrap()
    .with_input_bytes(u64::try_from(bytes.len()).unwrap())
    .unwrap();

    let file = tempfile::NamedTempFile::new().unwrap();
    fs::write(file.path(), &bytes).unwrap();
    let manager =
        MediaManager::new_with_limits_and_package_limits(file.path(), media_limits, package_limits)
            .unwrap();
    assert_eq!(manager.package_limits(), package_limits);

    let mut grown = bytes;
    grown.push(0);
    fs::write(file.path(), grown).unwrap();
    let error = manager.extract("image-7.png").unwrap_err();
    assert!(error.to_string().contains("iWork package input"));
}

#[cfg(unix)]
#[test]
fn directory_media_rejects_symbolic_links() -> std::io::Result<()> {
    use std::os::unix::fs::symlink;

    let bundle = tempfile::tempdir()?;
    let data = bundle.path().join("Data");
    fs::create_dir(&data)?;
    let outside = bundle.path().join("outside.png");
    fs::write(&outside, PNG)?;
    symlink(&outside, data.join("linked.png"))?;

    assert!(MediaManager::new(bundle.path()).is_err());

    let source_link = bundle.path().with_extension("-link");
    symlink(bundle.path(), &source_link)?;
    assert!(MediaManager::new(&source_link).is_err());
    Ok(())
}

#[test]
fn replaces_asset_and_preserves_unknown_metadata_fields() {
    let mut editor = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
    let before_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let before = &before_archive.object(2).unwrap().messages[0].data;
    let attributes_before = nested_field_bytes(before, 10);
    let outer_unknown_before = parse_wire_fields(before)
        .unwrap()
        .into_iter()
        .find(|field| field.number == 100)
        .map(|field| before[field.start..field.end].to_vec())
        .unwrap();

    let previous = editor.replace(asset_id(7), PNG).unwrap();
    assert_eq!(previous, b"\x89PNG\r\n\x1a\noriginal");
    assert_eq!(editor.extract(asset_id(7)).unwrap(), PNG);
    let asset = editor.asset(asset_id(7)).unwrap();
    assert_eq!(asset.digest, Sha1::digest(PNG).to_vec());
    assert_eq!(asset.declared_size, Some(PNG.len() as u64));
    assert_eq!(asset.message_reference_count, 1);

    let after_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let after = &after_archive.object(2).unwrap().messages[0].data;
    assert_eq!(nested_field_bytes(after, 10), attributes_before);
    let outer_unknown_after = parse_wire_fields(after)
        .unwrap()
        .into_iter()
        .find(|field| field.number == 100)
        .map(|field| after[field.start..field.end].to_vec())
        .unwrap();
    assert_eq!(outer_unknown_after, outer_unknown_before);
}

#[test]
fn rejects_type_mismatch_transactionally() {
    let mut editor = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
    let before = editor.to_bytes().unwrap();
    assert!(editor
        .replace(asset_id(7), b"%PDF-1.7\nnot-an-image")
        .is_err());
    assert_eq!(editor.to_bytes().unwrap(), before);
}

#[test]
fn removes_only_unreferenced_assets_transactionally() {
    let mut referenced = IWorkMediaEditor::from_package(synthetic_package()).unwrap();
    assert!(referenced.remove_unreferenced(asset_id(7)).is_err());
    assert!(referenced.asset(asset_id(7)).is_some());

    let mut package = synthetic_package();
    package
        .update_archive("Index/Document.iwa", |archive| {
            archive.objects[0].archive_info.message_infos[0]
                .data_references
                .clear();
            Ok(())
        })
        .unwrap();
    let before_archive = package.archive(PACKAGE_METADATA_ENTRY).unwrap();
    let before = &before_archive.object(2).unwrap().messages[0].data;
    let outer_unknown_before = parse_wire_fields(before)
        .unwrap()
        .into_iter()
        .find(|field| field.number == 100)
        .map(|field| before[field.start..field.end].to_vec())
        .unwrap();

    let mut editor = IWorkMediaEditor::from_package(package).unwrap();
    let removed = editor.remove_unreferenced(asset_id(7)).unwrap().unwrap();
    assert_eq!(removed, b"\x89PNG\r\n\x1a\noriginal");
    assert!(editor.asset(asset_id(7)).is_none());
    assert!(!editor.package().contains_entry("Data/image-7.png"));

    let after_archive = editor.package().archive(PACKAGE_METADATA_ENTRY).unwrap();
    let after = &after_archive.object(2).unwrap().messages[0].data;
    let outer_unknown_after = parse_wire_fields(after)
        .unwrap()
        .into_iter()
        .find(|field| field.number == 100)
        .map(|field| after[field.start..field.end].to_vec())
        .unwrap();
    assert_eq!(outer_unknown_after, outer_unknown_before);
}

#[test]
fn inserts_and_removes_unreferenced_asset_without_metadata_drift() {
    let package = synthetic_package();
    let initial_metadata = package
        .archive(PACKAGE_METADATA_ENTRY)
        .unwrap()
        .object(2)
        .unwrap()
        .messages[0]
        .data
        .clone();
    let mut editor = IWorkMediaEditor::from_package(package).unwrap();
    let inserted = editor.insert_unreferenced("new.png", PNG).unwrap();
    assert_eq!(inserted.data_identifier, asset_id(8));
    assert_eq!(inserted.package_path.as_deref(), Some("Data/new-8.png"));
    assert!(!inserted.is_referenced());
    assert_eq!(editor.extract(asset_id(8)).unwrap(), PNG);

    assert_eq!(editor.remove_unreferenced(asset_id(8)).unwrap().unwrap(), PNG);
    let final_metadata = editor
        .package()
        .archive(PACKAGE_METADATA_ENTRY)
        .unwrap()
        .object(2)
        .unwrap()
        .messages[0]
        .data
        .clone();
    assert_eq!(final_metadata, initial_metadata);
    assert!(editor.asset(asset_id(8)).is_none());
}

#[test]
fn wire_parser_rejects_truncation_and_groups() {
    assert!(parse_wire_fields(&[0x12, 0x05, 1]).is_err());
    assert!(parse_wire_fields(&[0x0b]).is_err());
    assert!(parse_wire_fields(&[0x80]).is_err());
}

#[test]
fn formats_sizes() {
    assert_eq!(format_bytes(0), "0.00 B");
    assert_eq!(format_bytes(1024), "1.00 KB");
    assert_eq!(format_bytes(1536 * 1024), "1.50 MB");
}
