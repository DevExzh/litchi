use super::worksheet::read;
use super::{Color, Colors, Formula, Group, Groups, Limits, Location, Sparkline, SparklineType};
use crate::raw::{Encoder, Kind, Records, kind};

fn stream(records: &[(Kind, &[u8])]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut writer = Encoder::new(&mut output, crate::raw::Limits::DEFAULT);
    for (record, payload) in records {
        writer.write_record(*record, payload).unwrap();
    }
    output
}

fn empty_sheet() -> Vec<u8> {
    stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::END_SHEET, &[]),
    ])
}

fn groups() -> Groups {
    let colors = Colors::uniform(Color::rgb(20, 40, 60, 255, 0));
    Groups::new(
        vec![
            Group::new(
                SparklineType::Line,
                colors,
                vec![Sparkline::new(Location::new(4, 7).unwrap(), None)],
                Limits::DEFAULT,
            )
            .unwrap(),
        ],
        Limits::DEFAULT,
    )
    .unwrap()
}

#[test]
fn absent_insert_uses_future_record_anchor_and_preserves_outside_bytes() {
    let unknown = Kind::new(0x1234).unwrap();
    let source = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (unknown, b"outside"),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (unknown, b"future"),
        (kind::FRT_END, &[]),
        (kind::END_SHEET, &[]),
    ]);
    let anchor = Records::new(&source, crate::raw::Limits::DEFAULT)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == kind::FRT_BEGIN)
        .unwrap()
        .offset();
    let mut edit = read(&source, Limits::DEFAULT).unwrap().edit();
    edit.set(groups());
    let commit = edit.commit().unwrap();
    let after = commit.patch().after();
    let mut block_end = None;
    let mut records = Records::new(after, crate::raw::Limits::DEFAULT);
    while let Some(record) = records.next() {
        if record.unwrap().kind() == kind::END_SPARKLINE_GROUPS {
            block_end = Some(records.offset());
            break;
        }
    }
    assert_eq!(&after[..anchor], &source[..anchor]);
    assert_eq!(&after[block_end.unwrap()..], &source[anchor..]);
    assert_eq!(commit.groups().unwrap(), &groups());
}

#[test]
fn rejects_sparkline_family_inside_or_after_the_frt_tail() {
    let block = super::encode_block(&groups(), Limits::DEFAULT).unwrap();
    let prefix = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
    ]);
    let close = stream(&[(kind::FRT_END, &[]), (kind::END_SHEET, &[])]);

    let inside = [&prefix[..], &block, &close].concat();
    assert!(read(&inside, Limits::DEFAULT).is_err());

    let closed_tail = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_END, &[]),
    ]);
    let end = stream(&[(kind::END_SHEET, &[])]);
    let after = [&closed_tail[..], &block, &end].concat();
    assert!(read(&after, Limits::DEFAULT).is_err());
}

#[test]
fn validates_nested_frt_tail_balance_and_preserves_balanced_tail() {
    let nested = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_END, &[]),
        (kind::FRT_END, &[]),
        (kind::END_SHEET, &[]),
    ]);
    let anchor = Records::new(&nested, crate::raw::Limits::DEFAULT)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == kind::FRT_BEGIN)
        .unwrap()
        .offset();
    let mut edit = read(&nested, Limits::DEFAULT).unwrap().edit();
    edit.set(groups());
    let commit = edit.commit().unwrap();
    let block_end = Records::new(commit.patch().after(), crate::raw::Limits::DEFAULT)
        .scan(0usize, |end, record| {
            let record = record.unwrap();
            *end = record.offset();
            Some((record.kind(), record.offset()))
        })
        .find_map(|(record, start)| (record == kind::FRT_BEGIN).then_some(start))
        .unwrap();
    assert_eq!(
        block_end,
        anchor
            + super::encode_block(&groups(), Limits::DEFAULT)
                .unwrap()
                .len()
    );
    assert_eq!(&commit.patch().after()[block_end..], &nested[anchor..]);

    let unclosed = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_END, &[]),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&unclosed, Limits::DEFAULT).is_err());

    let underflow = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_END, &[]),
        (kind::FRT_END, &[]),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&underflow, Limits::DEFAULT).is_err());
}

#[test]
fn refuses_absent_insertion_when_frt_top_level_cannot_be_proven() {
    let nested = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::BEGIN_CELL_WATCHES, &[]),
        (kind::FRT_BEGIN, &[0, 0, 0, 0]),
        (kind::FRT_END, &[]),
        (kind::END_CELL_WATCHES, &[]),
        (kind::END_SHEET, &[]),
    ]);
    let mut edit = read(&nested, Limits::DEFAULT).unwrap().edit();
    edit.set(groups());
    assert!(edit.commit().is_err());

    let mut no_frt = read(&empty_sheet(), Limits::DEFAULT).unwrap().edit();
    no_frt.set(groups());
    assert!(no_frt.commit().is_ok());
}

#[test]
fn existing_block_requires_provable_top_level_placement() {
    let block = super::encode_block(&groups(), Limits::DEFAULT).unwrap();
    let prefix = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
    ]);
    let end = stream(&[(kind::END_SHEET, &[])]);
    let direct = [&prefix[..], &block, &end].concat();
    assert_eq!(
        read(&direct, Limits::DEFAULT).unwrap().groups().unwrap(),
        &groups()
    );

    let nested_begin = stream(&[(kind::BEGIN_CELL_WATCHES, &[])]);
    let nested_end = stream(&[(kind::END_CELL_WATCHES, &[]), (kind::END_SHEET, &[])]);
    let nested = [&prefix[..], &nested_begin, &block, &nested_end].concat();
    assert!(read(&nested, Limits::DEFAULT).is_err());
}

#[test]
fn replace_remove_noop_and_inverse_are_byte_exact() {
    let source = empty_sheet();
    let base = read(&source, Limits::DEFAULT).unwrap();
    let noop = base.edit().commit().unwrap();
    assert!(noop.patch().is_empty());
    assert_eq!(noop.patch().after(), source);

    let mut insert = read(&source, Limits::DEFAULT).unwrap().edit();
    insert.set(groups());
    let inserted = insert.commit().unwrap();
    let inserted_bytes = inserted.patch().after().to_vec();
    let mut remove = read(&inserted_bytes, Limits::DEFAULT).unwrap().edit();
    assert!(remove.remove());
    let removed = remove.commit().unwrap();
    assert_eq!(removed.patch().after(), source);
    let inverse = inserted.into_patch().inverse();
    assert_eq!(inverse.apply(&inserted_bytes).unwrap().as_slice(), source);
}

#[test]
fn patch_refuses_stale_source_and_limits_cover_read_and_commit() {
    let source = empty_sheet();
    let mut edit = read(&source, Limits::DEFAULT).unwrap().edit();
    edit.set(groups());
    let commit = edit.commit().unwrap();
    let mut stale = source.clone();
    stale[0] ^= 1;
    assert!(commit.patch().apply(&stale).is_err());

    let tight = Limits::DEFAULT.with_block_bytes(1).unwrap();
    let mut edit = read(&source, Limits::DEFAULT).unwrap().edit_bounded(tight);
    edit.set(groups());
    assert!(edit.commit().is_err());
    assert!(read(commit.patch().after(), tight).is_err());
}

#[test]
fn worksheet_source_and_candidate_byte_limits_are_exact_and_atomic() {
    let source = empty_sheet();
    let source_exact = Limits::DEFAULT
        .with_block_bytes(source.len())
        .unwrap()
        .with_worksheet_bytes(source.len())
        .unwrap();
    assert!(read(&source, source_exact).is_ok());

    let source_over = Limits::DEFAULT
        .with_block_bytes(source.len() - 1)
        .unwrap()
        .with_worksheet_bytes(source.len() - 1)
        .unwrap();
    assert!(matches!(
        read(&source, source_over),
        Err(super::Error::Limit {
            resource: "worksheet bytes",
            actual,
            maximum,
        }) if actual == source.len() && maximum == source.len() - 1
    ));
    assert!(matches!(
        read(&source, Limits::DEFAULT)
            .unwrap()
            .edit_bounded(source_over)
            .commit(),
        Err(super::Error::Limit {
            resource: "worksheet bytes",
            actual,
            maximum,
        }) if actual == source.len() && maximum == source.len() - 1
    ));

    let encoded = super::encode_block(&groups(), Limits::DEFAULT).unwrap();
    let candidate_len = source.len() + encoded.len();
    let exact = Limits::DEFAULT
        .with_block_bytes(encoded.len())
        .unwrap()
        .with_worksheet_bytes(candidate_len)
        .unwrap();
    let mut edit = read(&source, exact).unwrap().edit();
    edit.set(groups());
    assert_eq!(edit.commit().unwrap().patch().after().len(), candidate_len);

    let over = exact.with_worksheet_bytes(candidate_len - 1).unwrap();
    let snapshot = read(&source, over).unwrap();
    let mut edit = snapshot.edit();
    edit.set(groups());
    assert!(matches!(
        edit.commit(),
        Err(super::Error::Limit {
            resource: "worksheet bytes",
            actual,
            maximum,
        }) if actual == candidate_len && maximum == candidate_len - 1
    ));
    assert_eq!(source, empty_sheet());
}

#[test]
fn strict_scan_rejects_family_members_outside_duplicate_unclosed_and_mismatch() {
    let outside = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::SPARKLINE, &[]),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&outside, Limits::DEFAULT).is_err());

    let unclosed = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::BEGIN_SPARKLINE_GROUPS, &[]),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&unclosed, Limits::DEFAULT).is_err());

    let mismatch = stream(&[
        (kind::BEGIN_SHEET, &[]),
        (kind::BEGIN_SHEET_DATA, &[]),
        (kind::END_SHEET_DATA, &[]),
        (kind::BEGIN_SPARKLINE_GROUPS, &[]),
        (kind::BEGIN_SPARKLINE_GROUP, &[]),
        (kind::END_SPARKLINE_GROUPS, &[]),
        (kind::END_SHEET, &[]),
    ]);
    assert!(read(&mismatch, Limits::DEFAULT).is_err());

    let block = super::encode_block(&groups(), Limits::DEFAULT).unwrap();
    let mut duplicate = Vec::new();
    let base = empty_sheet();
    let end = Records::new(&base, crate::raw::Limits::DEFAULT)
        .map(|record| record.unwrap())
        .find(|record| record.kind() == kind::END_SHEET)
        .unwrap()
        .offset();
    duplicate.extend_from_slice(&base[..end]);
    duplicate.extend_from_slice(&block);
    duplicate.extend_from_slice(&block);
    duplicate.extend_from_slice(&base[end..]);
    assert!(read(&duplicate, Limits::DEFAULT).is_err());
}

#[test]
fn workbook_api_applies_atomically_and_refuses_stale_commits() {
    let package = crate::Package::new().unwrap();
    let workbook_part = litchi_opc::PackURI::new("/xl/workbook.bin").unwrap();
    let workbook_before = package
        .opc_package()
        .get_part(&workbook_part)
        .unwrap()
        .blob()
        .to_vec();

    let mut edit = package.edit_sparklines().unwrap();
    assert!(edit.set(0, groups()).unwrap());
    let published = edit.commit().unwrap().into_package();
    let workbook = published.workbook().unwrap();
    assert_eq!(workbook.sparklines(0).unwrap().groups().unwrap(), &groups());
    assert_eq!(
        published
            .opc_package()
            .get_part(&workbook_part)
            .unwrap()
            .blob(),
        workbook_before
    );
    let reparsed = crate::Workbook::from_opc_package(published.opc_package().clone()).unwrap();
    assert_eq!(format!("{workbook:?}"), format!("{reparsed:?}"));

    let mut stale = published.edit_sparklines().unwrap();
    stale.remove(0).unwrap();
    let stale = stale.commit().unwrap();
    let mut other = published.edit_sparklines().unwrap();
    other.remove(0).unwrap();
    let other = other.commit().unwrap().into_package();
    // The patch checks the exact touched parts: a package whose worksheet
    // no longer holds the patch's before state refuses it...
    assert!(other.apply_sparklines(stale.patch()).is_err());
    // ...while one restored to that exact state accepts it.
    let mut again = other.edit_sparklines().unwrap();
    again.set(0, groups()).unwrap();
    let changed = again.commit().unwrap().into_package();
    assert!(changed.apply_sparklines(stale.patch()).is_ok());
}

#[test]
fn workbook_api_refuses_structural_formula_without_provable_context() {
    let package = crate::Package::new().unwrap();
    let colors = Colors::uniform(Color::rgb(20, 40, 60, 255, 0));
    let unbound = Groups::new(
        vec![
            Group::new(
                SparklineType::Line,
                colors,
                vec![Sparkline::new(
                    Location::new(4, 7).unwrap(),
                    Some(Formula::name(1, Limits::DEFAULT).unwrap()),
                )],
                Limits::DEFAULT,
            )
            .unwrap(),
        ],
        Limits::DEFAULT,
    )
    .unwrap();
    let before_bytes = package
        .workbook()
        .unwrap()
        .sparklines(0)
        .unwrap()
        .source_bytes()
        .to_vec();
    let mut edit = package.edit_sparklines().unwrap();
    assert!(edit.set(0, unbound).is_err());
    assert!(edit.commit().unwrap().patch().is_noop());
    assert_eq!(
        package
            .workbook()
            .unwrap()
            .sparklines(0)
            .unwrap()
            .source_bytes(),
        before_bytes
    );
}

#[test]
fn workbook_api_refuses_a_change_to_a_signed_workbook() {
    use litchi_opc::constants::{content_type, relationship_type};
    use litchi_opc::{BlobPart, PackURI};

    let mut opc = crate::Package::new().unwrap().into_opc();
    opc.try_add_part(Box::new(BlobPart::new(
        PackURI::new("/_xmlsignatures/origin.sigs").unwrap(),
        content_type::OPC_DIGITAL_SIGNATURE_ORIGIN.to_owned(),
        Vec::new(),
    )))
    .unwrap();
    opc.relate_to(
        "_xmlsignatures/origin.sigs",
        relationship_type::DIGITAL_SIGNATURE_ORIGIN,
    );
    let signed = crate::Package::from_opc(opc).unwrap();
    assert!(signed.is_signed());

    // A changed commit would invalidate the signature: refused, not stripped.
    let mut edit = signed.edit_sparklines().unwrap();
    let staged = edit.set(0, groups());
    assert!(matches!(
        staged.and_then(|_| edit.commit()),
        Err(crate::package::error::Error::Signed)
    ));
    assert!(signed.is_signed());

    // Removing the signature explicitly authorizes the edit.
    let unsigned = signed.unsign();
    let mut edit = unsigned.edit_sparklines().unwrap();
    edit.set(0, groups()).unwrap();
    let published = edit.commit().unwrap().into_package();
    assert_eq!(
        published
            .workbook()
            .unwrap()
            .sparklines(0)
            .unwrap()
            .groups()
            .unwrap(),
        &groups()
    );
    assert!(!published.is_signed());
}
