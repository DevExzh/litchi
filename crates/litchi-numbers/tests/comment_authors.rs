//! Comments in packages with more than one commenting author (record
//! 0329), built from `comment-metadata-native-resaved.numbers`: a new
//! comment is the generated author's, reused when the package holds it, else
//! a new generated author joins the author storage; no present author is
//! impersonated.

use litchi_iwa_archive::Limits;
use litchi_iwa_archive::iwa::{Archive, SnappyStream};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_numbers::{At, Package};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const AUTHOR: u32 = 212;
const COMMENT_STORAGE: u32 = 3_056;

fn bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let byte = (value & 0x7f) as u8;
        value >>= 7;
        if value == 0 {
            output.push(byte);
            return;
        }
        output.push(byte | 0x80);
    }
}

fn reference(identifier: u64) -> Vec<u8> {
    let mut payload = vec![0x08];
    varint(&mut payload, identifier);
    payload
}

fn field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    varint(&mut output, u64::from(number) << 3 | 2);
    varint(&mut output, payload.len() as u64);
    output.extend_from_slice(payload);
    output
}

/// Replace length-delimited field `number` (every occurrence) of `data`.
fn replace_field(data: &[u8], number: u32, payload: &[u8]) -> Vec<u8> {
    let view =
        litchi_iwa_common::wire::parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
            .expect("wire");
    let mut output = Vec::new();
    let mut done = false;
    for candidate in view.fields() {
        if candidate.number() == number {
            if !done {
                output.extend(field(number, payload));
                done = true;
            }
            continue;
        }
        output.extend_from_slice(candidate.raw());
    }
    if !done {
        output.extend(field(number, payload));
    }
    output
}

/// The varint field `number` of `data` set to `value`.
fn set_varint(data: &[u8], number: u32, value: u64) -> Vec<u8> {
    let view =
        litchi_iwa_common::wire::parse_wire_view(data, litchi_iwa_common::WireLimits::DEFAULT)
            .expect("wire");
    let mut output = Vec::new();
    for candidate in view.fields() {
        if candidate.number() == number {
            varint(&mut output, u64::from(number) << 3);
            varint(&mut output, value);
            continue;
        }
        output.extend_from_slice(candidate.raw());
    }
    output
}

/// The native comment fixture with a second comment by a second author:
/// the leaf writes the comment (reusing the package's one commenting
/// author, `litchi-iwa`), then the test points it at the package's other
/// author (listed in the author storage, commenting nowhere) and registers
/// the storage's edge to it. With `rename`, `litchi-iwa` is renamed too, so
/// the package holds no generated author.
fn two_authors(rename: bool) -> TestResult<(Package, At, At)> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/comment-metadata-native-resaved.numbers");
    let package = Package::open(path)?;
    let (native, _) = package
        .table_cell_comments(0, 0)?
        .into_iter()
        .next()
        .ok_or("a native comment")?;
    let semantic = package.sheets()[0].tables().next().ok_or("table")?;
    let added = At::new(semantic.row_count() - 1, semantic.column_count() - 1);
    let package = package
        .set_table_cell_comment(0, 0, added, "second")?
        .into_package();
    let source = bytes(&package)?;
    let catalog = Catalog::from_bytes(&source, Default::default())?;
    // The two authors, and the storages of the native and the new comment.
    let mut authors = Vec::new();
    let mut storages = Vec::new();
    let mut archives = Vec::new();
    for member in catalog
        .iter()
        .filter(|member| member.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(member.data(), Default::default())?;
        let archive = Archive::parse(stream.as_bytes(), Default::default())?;
        for object in &archive.objects {
            let identifier = object.archive_info.identifier.unwrap_or(0);
            match object.messages[0].type_ {
                AUTHOR => {
                    let name = litchi_iwa_common::wire::repeated_length_delimited_payloads(
                        &object.messages[0].data,
                        1,
                    )?
                    .first()
                    .map(|name| name.to_vec())
                    .unwrap_or_default();
                    authors.push((identifier, name));
                },
                COMMENT_STORAGE => storages.push(identifier),
                _ => {},
            }
        }
        archives.push((member.name().to_owned(), archive));
    }
    let generated = authors
        .iter()
        .find(|(_, name)| name == b"litchi-iwa")
        .ok_or("generated author")?
        .0;
    let other = authors
        .iter()
        .find(|(_, name)| name != b"litchi-iwa")
        .ok_or("second author")?
        .0;
    let latest = *storages.iter().max().ok_or("storages")?;
    let mut storage_component = None;
    let mut edits = Vec::new();
    for (name, mut archive) in archives {
        let mut changed = false;
        for object in &mut archive.objects {
            let kind = object.messages[0].type_;
            if kind == AUTHOR && rename && object.archive_info.identifier == Some(generated) {
                object.messages[0].data = replace_field(&object.messages[0].data, 1, b"Other");
                changed = true;
            }
            if kind == COMMENT_STORAGE && object.archive_info.identifier == Some(latest) {
                object.messages[0].data =
                    replace_field(&object.messages[0].data, 3, &reference(other));
                let info = &mut object.archive_info.message_infos[0];
                for list in std::iter::once(&mut info.object_references).chain(
                    info.field_infos
                        .iter_mut()
                        .map(|field_info| &mut field_info.object_references),
                ) {
                    for identifier in list.iter_mut() {
                        if *identifier == generated {
                            *identifier = other;
                        }
                    }
                }
                storage_component = Some(name.clone());
                changed = true;
            }
        }
        if changed {
            edits.push((name, archive));
        }
    }
    // The registry: the storage's component reads the other author.
    let storage_component = storage_component.ok_or("storage component")?;
    let locator = storage_component
        .trim_start_matches("Index/")
        .trim_end_matches(".iwa")
        .to_owned();
    let metadata = catalog
        .iter()
        .find(|member| member.name() == "Index/Metadata.iwa")
        .ok_or("metadata")?;
    let mut registry = Archive::parse(
        SnappyStream::decompress(metadata.data(), Default::default())?.as_bytes(),
        Default::default(),
    )?;
    for object in &mut registry.objects {
        for message in &mut object.messages {
            if message.type_ != 11_006 {
                continue;
            }
            let view = litchi_iwa_common::wire::parse_wire_view(
                &message.data,
                litchi_iwa_common::WireLimits::DEFAULT,
            )?;
            let mut output = Vec::new();
            for candidate in view.fields() {
                if candidate.number() != 3 {
                    output.extend_from_slice(candidate.raw());
                    continue;
                }
                let component = candidate.payload();
                let names_it =
                    litchi_iwa_common::wire::repeated_length_delimited_payloads(component, 3)?
                        .first()
                        .is_some_and(|name| *name == locator.as_bytes());
                if !names_it {
                    output.extend_from_slice(candidate.raw());
                    continue;
                }
                let edge =
                    litchi_iwa_common::wire::repeated_length_delimited_payloads(component, 6)?
                        .first()
                        .map(|edge| set_varint(edge, 2, other))
                        .ok_or("edge")?;
                let mut rewritten = component.to_vec();
                rewritten.extend(field(6, &edge));
                output.extend(field(3, &rewritten));
            }
            message.data = output;
        }
    }
    edits.push(("Index/Metadata.iwa".to_owned(), registry));
    let compressed: Vec<(String, Vec<u8>)> = edits
        .into_iter()
        .map(|(name, archive)| -> TestResult<(String, Vec<u8>)> {
            Ok((
                name,
                SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?,
            ))
        })
        .collect::<TestResult<_>>()?;
    let entries: Vec<EntryEdit<'_>> = compressed
        .iter()
        .map(|(name, data)| EntryEdit::new(name, data))
        .collect();
    let package = Package::from_bytes(&catalog.reassemble_to_bytes(&entries, Limits::default())?)?;
    Ok((package, native, added))
}

fn author(package: &Package, position: At) -> Option<String> {
    package
        .table_cell_comment(0, 0, position)
        .ok()
        .flatten()
        .and_then(|comment| {
            comment
                .author()
                .and_then(|author| author.display_name().map(str::to_owned))
        })
}

#[test]
fn the_generated_author_is_reused_among_several() -> TestResult {
    let (package, native, added) = two_authors(false)?;
    assert_eq!(author(&package, native).as_deref(), Some("litchi-iwa"));
    assert_ne!(author(&package, added).as_deref(), Some("litchi-iwa"));
    let source = bytes(&package)?;
    let third = At::new(added.row(), added.column() - 1);
    let commit = package.set_table_cell_comment(0, 0, third, "third")?;
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    assert_eq!(author(&reopened, third).as_deref(), Some("litchi-iwa"));
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn a_new_generated_author_joins_the_storage() -> TestResult {
    let (package, native, added) = two_authors(true)?;
    assert_eq!(author(&package, native).as_deref(), Some("Other"));
    let present = author(&package, added);
    let source = bytes(&package)?;
    let third = At::new(added.row(), added.column() - 1);
    let commit = package.set_table_cell_comment(0, 0, third, "third")?;
    let reopened = Package::from_bytes(&bytes(commit.package())?)?;
    assert_eq!(author(&reopened, third).as_deref(), Some("litchi-iwa"));
    // The present authors keep their comments.
    assert_eq!(author(&reopened, native).as_deref(), Some("Other"));
    assert_eq!(author(&reopened, added), present);
    let restored = commit
        .package()
        .apply_table_cell_comment(&commit.patch().inverse())?;
    assert_eq!(bytes(restored.package())?, source);
    // A further comment reuses the new generated author.
    let fourth = At::new(added.row() - 1, added.column());
    let again = reopened.set_table_cell_comment(0, 0, fourth, "fourth")?;
    assert_eq!(
        author(again.package(), fourth).as_deref(),
        Some("litchi-iwa")
    );
    Ok(())
}
