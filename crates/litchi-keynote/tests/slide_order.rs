use std::io;
use std::sync::Arc;

use litchi_iwa_archive::package::Catalog;
use litchi_iwa_common::wire::{WireView, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsk, tsp};
use litchi_keynote::{
    Limits, Package, Position, ReadError, ReadOptions, SemanticLimits, SlideOrderError,
    SlideOrderLimitKind, SlideSelector,
};
use prost::Message as _;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const NODE_IDS: [u64; 4] = [3, 5, 7, 9];
const SLIDE_IDS: [u64; 4] = [4, 6, 8, 10];
const NAMES: [&str; 4] = ["Intro", "Plan", "Evidence", "Appendix"];

type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, Copy)]
enum Topology {
    Flat,
    DuplicateNode,
    AliasedSlide,
    DeprecatedRoot,
    SecondaryList,
    ChildNode,
    DeepNode,
    ZeroNode,
    ZeroSlide,
    MissingSlide,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        deprecated_type: Some(-1),
        deprecated_is_external: Some(false),
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

#[allow(
    clippy::cast_possible_truncation,
    reason = "each emitted byte intentionally retains only the low seven varint bits"
)]
fn push_varint(mut value: u64, output: &mut Vec<u8>) {
    while value >= 0x80 {
        output.push(((value as u8) & 0x7f) | 0x80);
        value >>= 7;
    }
    output.push(value as u8);
}

fn length_delimited_field(number: u32, payload: &[u8]) -> Vec<u8> {
    let mut field = Vec::with_capacity(payload.len().saturating_add(8));
    push_varint((u64::from(number) << 3) | 2, &mut field);
    push_varint(payload.len() as u64, &mut field);
    field.extend_from_slice(payload);
    field
}

fn fixed32_field(number: u32, value: u32) -> Vec<u8> {
    let mut field = Vec::with_capacity(8);
    push_varint((u64::from(number) << 3) | 5, &mut field);
    field.extend_from_slice(&value.to_le_bytes());
    field
}

fn fixed64_field(number: u32, value: u64) -> Vec<u8> {
    let mut field = Vec::with_capacity(12);
    push_varint((u64::from(number) << 3) | 1, &mut field);
    field.extend_from_slice(&value.to_le_bytes());
    field
}

fn adversarial_reference(identifier_value: u64, sentinel: u64) -> TestResult<Vec<u8>> {
    let canonical = reference(identifier_value).encode_to_vec();
    let view = WireView::parse(&canonical)?;
    let identifier_field = view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("reference identifier is missing"))?;
    let deprecated_type = view
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("reference deprecated type is missing"))?;
    let external = view
        .fields()
        .find(|field| field.number() == 3)
        .ok_or_else(|| io::Error::other("reference external flag is missing"))?;

    let mut payload = Vec::new();
    append_varint_field(&mut payload, 99, sentinel)?;
    payload.extend_from_slice(external.raw());
    let long_unknown = [b'x'; 130];
    let unknown = if sentinel == 16_384 {
        long_unknown.as_slice()
    } else {
        b"unknown-reference".as_slice()
    };
    payload.extend_from_slice(&length_delimited_field(100, unknown));
    payload.extend_from_slice(identifier_field.raw());
    payload.extend_from_slice(&fixed64_field(101, 0x1122_3344_5566_7788));
    payload.extend_from_slice(deprecated_type.raw());
    payload.extend_from_slice(&fixed32_field(102, 0xaabb_ccdd));
    Ok(payload)
}

fn raw_show(topology: Topology, noncanonical_mode: bool) -> TestResult<Vec<u8>> {
    let node_ids = match topology {
        Topology::DuplicateNode => [3, 3, 7, 9],
        Topology::ZeroNode => [0, 5, 7, 9],
        Topology::Flat
        | Topology::AliasedSlide
        | Topology::DeprecatedRoot
        | Topology::SecondaryList
        | Topology::ChildNode
        | Topology::DeepNode
        | Topology::ZeroSlide
        | Topology::MissingSlide => NODE_IDS,
    };
    let mut tree = Vec::new();
    if matches!(topology, Topology::DeprecatedRoot) {
        tree.extend_from_slice(&length_delimited_field(1, &reference(3).encode_to_vec()));
    }
    for (index, identifier) in node_ids.into_iter().enumerate() {
        append_varint_field(
            &mut tree,
            70 + u32::try_from(index)?,
            u64::try_from(index)?.saturating_add(31),
        )?;
        let payload = adversarial_reference(
            identifier,
            if index == 3 {
                16_384
            } else {
                u64::try_from(index)?.saturating_add(71)
            },
        )?;
        tree.extend_from_slice(&length_delimited_field(2, &payload));
    }
    let canonical = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive::default(),
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        mode: Some(-1),
        slide_list: matches!(topology, Topology::SecondaryList).then(|| reference(82)),
        ..Default::default()
    }
    .encode_to_vec();
    let view = WireView::parse(&canonical)?;
    let mut show = Vec::with_capacity(canonical.len().saturating_add(tree.len()));
    for field in view.fields() {
        match field.number() {
            3 => {
                show.extend_from_slice(&fixed32_field(100, 0x1020_3040));
                show.extend_from_slice(&length_delimited_field(3, &tree));
                show.extend_from_slice(&length_delimited_field(101, b"unknown-show"));
            },
            9 if noncanonical_mode => {
                // A five-byte unsigned alias for -1 is not canonical protobuf
                // int32; negative int32 values require ten-byte sign extension.
                show.extend_from_slice(&[0x48, 0xff, 0xff, 0xff, 0xff, 0x0f]);
            },
            _ => show.extend_from_slice(field.raw()),
        }
    }
    Ok(show)
}

#[allow(
    deprecated,
    reason = "native schema retains required legacy cache fields"
)]
fn node(index: usize, topology: Topology) -> Vec<u8> {
    let slide_identifier = if matches!(topology, Topology::AliasedSlide) && index == 1 {
        SLIDE_IDS[0]
    } else if matches!(topology, Topology::ZeroSlide) && index == 0 {
        0
    } else if matches!(topology, Topology::MissingSlide) && index == 0 {
        123
    } else {
        SLIDE_IDS[index]
    };
    kn::SlideNodeArchive {
        slide: Some(reference(slide_identifier)),
        children: if matches!(topology, Topology::ChildNode) && index == 0 {
            vec![reference(NODE_IDS[1])]
        } else {
            Vec::new()
        },
        depth: (matches!(topology, Topology::DeepNode) && index == 0).then_some(2),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..Default::default()
    }
    .encode_to_vec()
}

fn slide(name: &str) -> Vec<u8> {
    kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        name: Some(name.to_owned()),
        in_document: true,
        ..Default::default()
    }
    .encode_to_vec()
}

fn package_bytes(
    topology: Topology,
    names: [&str; 4],
    noncanonical_mode: bool,
) -> TestResult<Vec<u8>> {
    package_bytes_with_iwa_framing(topology, names, noncanonical_mode, false)
}

fn package_bytes_with_iwa_framing(
    topology: Topology,
    names: [&str; 4],
    noncanonical_mode: bool,
    noncanonical_iwa_prefix: bool,
) -> TestResult<Vec<u8>> {
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..Default::default()
        },
        show: reference(2),
        ..Default::default()
    };
    let show_object = ArchiveObject::new(
        2,
        vec![
            RawMessage {
                type_: 777,
                data: b"before-show-sentinel".to_vec(),
            },
            RawMessage {
                type_: 2,
                data: raw_show(topology, noncanonical_mode)?,
            },
            RawMessage {
                type_: 778,
                data: b"after-show-sentinel".to_vec(),
            },
        ],
    )?;
    let mut document_objects = vec![object(1, 1, document.encode_to_vec())?, show_object];
    for (index, identifier) in NODE_IDS.into_iter().enumerate() {
        document_objects.push(object(identifier, 4, node(index, topology))?);
    }
    if matches!(topology, Topology::ZeroNode) {
        document_objects.push(object(0, 4, node(0, topology))?);
    }
    let document_component = if noncanonical_iwa_prefix {
        component_with_noncanonical_object_prefix(document_objects)?
    } else {
        component(document_objects)?
    };

    let mut entries: Vec<(String, Vec<u8>)> = vec![(
        "Data/sentinel.bin".to_owned(),
        b"unrelated opaque sentinel".to_vec(),
    )];
    for ((identifier, name), index) in SLIDE_IDS.into_iter().zip(names).zip(0usize..) {
        entries.push((
            format!("Index/Slide-{identifier}.iwa"),
            component(vec![object(identifier, 5, slide(name))?])?,
        ));
        assert!(index < 4);
    }
    if matches!(topology, Topology::ZeroSlide) {
        entries.push((
            "Index/Slide-0.iwa".to_owned(),
            component(vec![object(0, 5, slide(names[0]))?])?,
        ));
    }
    entries.push((DOCUMENT_MEMBER.to_owned(), document_component));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn component_with_noncanonical_object_prefix(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    let canonical = Archive { objects }.to_bytes()?;
    let (header_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(&canonical)?;
    if prefix_length != 1 || header_length >= 0x80 {
        return Err(io::Error::other("fixture expected a one-byte object prefix").into());
    }
    let mut noncanonical = Vec::with_capacity(canonical.len().saturating_add(1));
    noncanonical.push(u8::try_from(header_length)? | 0x80);
    noncanonical.push(0);
    noncanonical.extend_from_slice(&canonical[prefix_length..]);
    Archive::parse(&noncanonical)?;
    Ok(SnappyStream::compress(&noncanonical)?)
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner =
        litchi_iwa_archive::package::to_bytes(inner_entries.iter().copied(), Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.key/Index.zip", inner.as_slice()),
            (
                "legacy.key/Data/sentinel.bin",
                b"legacy outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn names(package: &Package) -> TestResult<Vec<String>> {
    Ok(package
        .show()?
        .slides()
        .iter()
        .map(|slide| slide.name().unwrap_or_default().to_owned())
        .collect())
}

fn document_stream(package: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("missing synthetic document member"))?;
    Ok(SnappyStream::decompress(entry.data())?.into_bytes())
}

fn slide_reference_records(package: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(slide_tree_records(package)?
        .into_iter()
        .filter_map(|(number, raw)| (number == 2).then_some(raw))
        .collect())
}

fn slide_tree_records(package: &[u8]) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let show = archive
        .object(2)
        .and_then(|object| object.messages.iter().find(|message| message.type_ == 2))
        .ok_or_else(|| io::Error::other("missing synthetic show"))?;
    let show_view = WireView::parse(&show.data)?;
    let tree = show_view
        .fields()
        .find(|field| field.number() == 3)
        .ok_or_else(|| io::Error::other("missing synthetic slide tree"))?;
    Ok(WireView::parse(tree.payload())?
        .fields()
        .map(|field| (field.number(), field.raw().to_vec()))
        .collect())
}

fn show_records(package: &[u8]) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let stream = document_stream(package)?;
    let archive = Archive::parse(&stream)?;
    let show = archive
        .object(2)
        .and_then(|object| object.messages.iter().find(|message| message.type_ == 2))
        .ok_or_else(|| io::Error::other("missing synthetic show"))?;
    Ok(WireView::parse(&show.data)?
        .fields()
        .map(|field| (field.number(), field.raw().to_vec()))
        .collect())
}

fn show_message_range(stream: &[u8]) -> TestResult<std::ops::Range<usize>> {
    let archive = Archive::parse(stream)?;
    let object = archive
        .object(2)
        .ok_or_else(|| io::Error::other("missing synthetic show object"))?;
    let (message_index, message) = object
        .messages
        .iter()
        .enumerate()
        .find(|(_index, message)| message.type_ == 2)
        .ok_or_else(|| io::Error::other("missing synthetic show message"))?;
    let preceding = object
        .messages
        .iter()
        .take(message_index)
        .map(|preceding_message| preceding_message.data.len())
        .sum::<usize>();
    let start = usize::try_from(object.data_offset)? + preceding;
    Ok(start..start + message.data.len())
}

#[test]
fn every_four_slide_move_matches_final_position_model() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    for source in 0..4 {
        for destination in 0..4 {
            let package = Package::from_bytes(&bytes)?;
            let mut expected = NAMES.map(str::to_owned).to_vec();
            let moved = expected.remove(source);
            expected.insert(destination, moved);

            let mut edit = package.edit_slide_order();
            edit.move_slide(Position::new(source), Position::new(destination))?;
            let commit = edit.commit()?;
            assert_eq!(names(commit.package())?, expected);
            assert_eq!(names(&package)?, NAMES);
            assert_eq!(commit.patch().source_position(), Position::new(source));
            assert_eq!(
                commit.patch().destination_position(),
                Position::new(destination)
            );
            assert_eq!(commit.diagnostics().changed(), source != destination);
            assert_eq!(
                commit.diagnostics().touched_components(),
                usize::from(source != destination)
            );
            assert_eq!(
                commit.diagnostics().full_reparse_performed(),
                source != destination
            );
        }
    }
    Ok(())
}

#[test]
fn selectors_validate_source_before_final_destination_and_bound_one_operation() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let package = Package::from_bytes(&bytes)?;

    let mut by_name = package.edit_slide_order();
    by_name.move_slide("Appendix", Position::new(0))?;
    assert_eq!(
        names(by_name.commit()?.package())?,
        ["Appendix", "Intro", "Plan", "Evidence"]
    );

    let mut missing_position = package.edit_slide_order();
    assert!(matches!(
        missing_position.move_slide(Position::new(99), Position::new(99)),
        Err(SlideOrderError::SlidePositionNotFound { position }) if position == Position::new(99)
    ));
    let mut missing_name = package.edit_slide_order();
    assert!(matches!(
        missing_name.move_slide("Missing", Position::new(99)),
        Err(SlideOrderError::SlideNameNotFound)
    ));
    let mut bad_destination = package.edit_slide_order();
    assert!(matches!(
        bad_destination.move_slide(0usize, Position::new(4)),
        Err(SlideOrderError::DestinationOutOfRange { position, slide_count: 4 })
            if position == Position::new(4)
    ));
    let mut one = package.edit_slide_order();
    one.move_slide(0usize, Position::new(1))?;
    assert!(matches!(
        one.move_slide(1usize, Position::new(0)),
        Err(SlideOrderError::OperationAlreadyStaged)
    ));
    assert!(matches!(
        package.edit_slide_order().commit(),
        Err(SlideOrderError::NoStagedOperation)
    ));

    let ambiguous_bytes = package_bytes(Topology::Flat, ["Same", "Same", "C", "D"], false)?;
    let ambiguous = Package::from_bytes(&ambiguous_bytes)?;
    let mut edit = ambiguous.edit_slide_order();
    assert!(matches!(
        edit.move_slide(SlideSelector::name("Same"), Position::new(0)),
        Err(SlideOrderError::AmbiguousSelector)
    ));
    Ok(())
}

#[test]
fn noop_reuses_source_and_changed_patch_is_exact_and_reversible() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let mut noop_edit = package.edit_slide_order();
    noop_edit.move_slide(2usize, Position::new(2))?;
    let noop_commit = noop_edit.commit()?;
    assert_eq!(
        noop_commit.package().source_bytes().as_ptr(),
        source_pointer
    );
    assert_eq!(noop_commit.package().source_bytes(), bytes);
    assert!(noop_commit.patch().is_noop());

    let mut changed_edit = package.edit_slide_order();
    changed_edit.move_slide("Appendix", Position::new(0))?;
    let changed_commit = changed_edit.commit()?;
    let applied = package.apply_slide_order(changed_commit.patch())?;
    assert_eq!(
        applied.package().source_bytes(),
        changed_commit.package().source_bytes()
    );
    let inverse = changed_commit.patch().inverse();
    assert_eq!(inverse.inverse(), changed_commit.patch().clone());
    let restored = changed_commit.package().apply_slide_order(&inverse)?;
    assert_eq!(restored.package().source_bytes(), bytes);

    let unrelated_bytes = package_bytes(
        Topology::Flat,
        ["Other", "Plan", "Evidence", "Appendix"],
        false,
    )?;
    let unrelated = Package::from_bytes(&unrelated_bytes)?;
    assert!(matches!(
        unrelated.apply_slide_order(changed_commit.patch()),
        Err(SlideOrderError::PatchConflict)
    ));
    let debug = format!("{:?}", changed_commit.patch());
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("Appendix"));
    assert!(!debug.contains("fingerprint"));
    assert!(!debug.contains("bytes"));
    Ok(())
}

#[test]
fn every_commit_validates_downstream_semantics_before_publication() -> TestResult<()> {
    let bytes = package_bytes(Topology::MissingSlide, NAMES, false)?;
    let package = Package::from_bytes(&bytes)?;
    let source_pointer = package.source_bytes().as_ptr();
    let mut edit = package.edit_slide_order();
    edit.move_slide(0usize, Position::new(0))?;
    assert!(matches!(edit.commit(), Err(SlideOrderError::InvalidSource)));

    let mut changed = package.edit_slide_order();
    changed.move_slide(0usize, Position::new(1))?;
    assert!(matches!(
        changed.commit(),
        Err(SlideOrderError::InvalidSource)
    ));
    assert_eq!(package.source_bytes().as_ptr(), source_pointer);
    assert_eq!(package.source_bytes(), bytes);
    Ok(())
}

#[test]
fn rewrite_permutates_complete_reference_records_and_only_one_component() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let package = Package::from_bytes(&bytes)?;
    let before_records = slide_reference_records(&bytes)?;
    let before_tree = slide_tree_records(&bytes)?;
    let before_show = show_records(&bytes)?;
    let before_stream = document_stream(&bytes)?;
    let mut edit = package.edit_slide_order();
    edit.move_slide(3usize, Position::new(1))?;
    let commit = edit.commit()?;
    let after_bytes = commit.package().source_bytes();
    let after_records = slide_reference_records(after_bytes)?;
    let after_tree = slide_tree_records(after_bytes)?;
    let after_show = show_records(after_bytes)?;
    assert_eq!(
        after_records,
        [
            before_records[0].clone(),
            before_records[3].clone(),
            before_records[1].clone(),
            before_records[2].clone(),
        ]
    );
    let mut expected_tree = before_tree.clone();
    let moved_records = [
        before_records[0].clone(),
        before_records[3].clone(),
        before_records[1].clone(),
        before_records[2].clone(),
    ];
    let mut moved_index = 0usize;
    for (number, raw) in &mut expected_tree {
        if *number == 2 {
            *raw = moved_records[moved_index].clone();
            moved_index += 1;
        }
    }
    assert_eq!(after_tree, expected_tree);
    assert_eq!(after_show.len(), before_show.len());
    for (before_record, after_record) in before_show.iter().zip(&after_show) {
        assert_eq!(before_record.0, after_record.0);
        if before_record.0 == 3 {
            assert_eq!(before_record.1.len(), after_record.1.len());
            assert_ne!(before_record.1, after_record.1);
        } else {
            assert_eq!(before_record.1, after_record.1);
        }
    }
    assert_eq!(document_stream(after_bytes)?.len(), before_stream.len());

    let before_catalog = Catalog::from_bytes(&bytes)?;
    let after_catalog = Catalog::from_bytes(after_bytes)?;
    let mut changed_entries = 0usize;
    for (before_entry, after_entry) in before_catalog.iter().zip(after_catalog.iter()) {
        assert_eq!(before_entry.name(), after_entry.name());
        if before_entry.data() == after_entry.data() {
            assert_eq!(
                before_entry.raw_record().local_record(),
                after_entry.raw_record().local_record()
            );
            assert_eq!(
                before_entry.raw_record().central_directory_record(),
                after_entry.raw_record().central_directory_record()
            );
        } else {
            changed_entries += 1;
            assert_eq!(before_entry.name(), DOCUMENT_MEMBER);
        }
    }
    assert_eq!(changed_entries, 1);

    let mut reverse = commit.package().edit_slide_order();
    reverse.move_slide(1usize, Position::new(3))?;
    let reversed = reverse.commit()?;
    assert_eq!(
        document_stream(reversed.package().source_bytes())?,
        before_stream
    );
    Ok(())
}

#[test]
fn mutation_rejects_nonflat_or_aliased_identity_but_reader_remains_compatible() -> TestResult<()> {
    for topology in [
        Topology::DuplicateNode,
        Topology::AliasedSlide,
        Topology::DeprecatedRoot,
        Topology::SecondaryList,
        Topology::ChildNode,
        Topology::DeepNode,
        Topology::ZeroNode,
        Topology::ZeroSlide,
    ] {
        let bytes = package_bytes(topology, NAMES, false)?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(package.show()?.slides().len(), 4, "{topology:?}");

        let mut changed = package.edit_slide_order();
        changed.move_slide(0usize, Position::new(1))?;
        assert!(
            matches!(changed.commit(), Err(SlideOrderError::UnsupportedTopology)),
            "{topology:?}"
        );

        let mut noop = package.edit_slide_order();
        noop.move_slide(0usize, Position::new(0))?;
        assert!(noop.commit()?.patch().is_noop(), "{topology:?}");
    }
    Ok(())
}

#[test]
fn changed_rewrite_preserves_noncanonical_iwa_prefix_and_all_bytes_outside_show() -> TestResult<()>
{
    let bytes = package_bytes_with_iwa_framing(Topology::Flat, NAMES, false, true)?;
    let package = Package::from_bytes(&bytes)?;
    let before = document_stream(&bytes)?;
    let range = show_message_range(&before)?;
    assert!(before.starts_with(&[before[0] | 0x80, 0]));

    let mut edit = package.edit_slide_order();
    edit.move_slide(0usize, Position::new(3))?;
    let commit = edit.commit()?;
    let after = document_stream(commit.package().source_bytes())?;
    assert_eq!(&after[..range.start], &before[..range.start]);
    assert_eq!(&after[range.end..], &before[range.end..]);
    assert_ne!(&after[range.clone()], &before[range]);
    Ok(())
}

#[test]
fn legacy_physical_source_allows_exact_noop_but_refuses_changed_reassembly() -> TestResult<()> {
    let flat = package_bytes(Topology::Flat, NAMES, false)?;
    let legacy = legacy_package_bytes(&flat)?;
    let package = Package::from_bytes(&legacy)?;
    assert_eq!(names(&package)?, NAMES);

    let mut noop_edit = package.edit_slide_order();
    noop_edit.move_slide(1usize, Position::new(1))?;
    let noop_commit = noop_edit.commit()?;
    assert_eq!(noop_commit.package().source_bytes(), legacy);
    assert!(noop_commit.patch().is_noop());
    let applied = package.apply_slide_order(noop_commit.patch())?;
    assert_eq!(applied.package().source_bytes(), legacy);

    let mut changed = package.edit_slide_order();
    changed.move_slide(0usize, Position::new(1))?;
    assert!(matches!(
        changed.commit(),
        Err(SlideOrderError::UnsupportedSource)
    ));
    Ok(())
}

#[test]
fn signed_int32_is_strict_and_commit_retains_read_options() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let options = ReadOptions::default();
    let package = Package::from_bytes_with_options(&bytes, options)?;
    assert_eq!(package.read_options(), options);
    package.validate()?;
    let mut edit = package.edit_slide_order();
    edit.move_slide(0usize, Position::new(3))?;
    let commit = edit.commit()?;
    assert_eq!(commit.package().read_options(), options);

    let bad_bytes = package_bytes(Topology::Flat, NAMES, true)?;
    let bad_package = Package::from_bytes(&bad_bytes)?;
    assert!(matches!(
        bad_package.show(),
        Err(ReadError::InvalidFormat(_) | ReadError::Decode(_))
    ));
    Ok(())
}

#[test]
fn slide_limit_is_inclusive_and_typed_for_transaction_staging() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let semantic_limits = |slides| {
        SemanticLimits::new(
            SemanticLimits::MAX_OBJECTS,
            slides,
            SemanticLimits::MAX_REFERENCES,
            SemanticLimits::MAX_TEXT_STORAGES,
            SemanticLimits::MAX_TEXT_FRAGMENTS,
            SemanticLimits::MAX_TEXT_BYTES,
        )
    };

    let exact = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), semantic_limits(NAMES.len())?),
    )?;
    let mut exact_edit = exact.edit_slide_order();
    exact_edit.move_slide(0usize, Position::new(NAMES.len() - 1))?;
    assert_eq!(
        names(exact_edit.commit()?.package())?,
        ["Plan", "Evidence", "Appendix", "Intro"]
    );

    let one_under = Package::from_bytes_with_options(
        &bytes,
        ReadOptions::new(Limits::default(), semantic_limits(NAMES.len() - 1)?),
    )?;
    let mut limited_edit = one_under.edit_slide_order();
    assert!(matches!(
        limited_edit.move_slide(0usize, Position::new(1)),
        Err(SlideOrderError::LimitExceeded {
            kind: SlideOrderLimitKind::Slides,
            observed,
            maximum,
        }) if observed == NAMES.len() as u64 && maximum == (NAMES.len() - 1) as u64
    ));
    Ok(())
}

#[test]
fn concurrent_edits_publish_independent_immutable_snapshots() -> TestResult<()> {
    let bytes = package_bytes(Topology::Flat, NAMES, false)?;
    let package = Package::from_bytes(&bytes)?;
    let forward_source = package.clone();
    let backward_source = package.clone();

    let forward_worker = std::thread::spawn(move || {
        let mut edit = forward_source.edit_slide_order();
        edit.move_slide(0usize, Position::new(3))?;
        edit.commit()
    });
    let backward_worker = std::thread::spawn(move || {
        let mut edit = backward_source.edit_slide_order();
        edit.move_slide(3usize, Position::new(0))?;
        edit.commit()
    });
    let forward = forward_worker
        .join()
        .map_err(|_panic| io::Error::other("forward slide-order worker panicked"))??;
    let backward = backward_worker
        .join()
        .map_err(|_panic| io::Error::other("backward slide-order worker panicked"))??;

    assert_eq!(
        names(forward.package())?,
        ["Plan", "Evidence", "Appendix", "Intro"]
    );
    assert_eq!(
        names(backward.package())?,
        ["Appendix", "Intro", "Plan", "Evidence"]
    );
    assert_eq!(package.source_bytes(), bytes);
    assert_ne!(
        forward.package().source_bytes(),
        backward.package().source_bytes()
    );
    Ok(())
}

#[test]
fn public_slide_order_values_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<litchi_keynote::SlideOrderCommit>();
    assert_send_sync::<litchi_keynote::SlideOrderPatch>();
    assert_send_sync::<litchi_keynote::SlideOrderDiagnostics>();
    assert_send_sync::<SlideOrderError>();
    assert_send_sync::<Arc<[u8]>>();
}
