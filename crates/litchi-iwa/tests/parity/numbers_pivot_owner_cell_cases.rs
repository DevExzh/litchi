//! Cell edits on a Numbers table that owns pivot or active category state
//! (record 0424a): the host's formula setters against the leaf's
//! `edit_table_cells` on the same fixtures. The host writes cells of such
//! tables (it never checks their pivot or category owners), so the leaf no
//! longer refuses them.
//!
//! `basic.numbers`' first table natively names a category owner (field 86,
//! a `TST.CategoryOwnerRefArchive` whose group-by is disabled) and a
//! deprecated inline category owner (field 81, disabled). The fixtures
//! enable the referenced group-by, enable the inline one, or give the
//! model a pivot owner (field 85, a `TST.PivotOwnerArchive`).

use litchi_iwa::numbers;
use litchi_iwa_archive::iwa::{
    Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream,
};
use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::formula::{
    FormulaBinaryOperator, FormulaCachedValue, FormulaCellReference, FormulaExpression,
};
use litchi_iwa_protos::{tsp, tst};
use litchi_numbers::formula::{BinaryOperator, CachedValue, CellReference, Expression};
use litchi_numbers::table::At;
use prost::Message;

use crate::harness::{Case, Evidence, Fixture, Format, LeafRun, TestResult, Tier, fixture};
use crate::leaf_bytes;

const ENGINE: &str = "Index/CalculationEngine.iwa";
/// `basic.numbers`' first table model, its category owner's group-by, and
/// a new pivot owner.
const MODEL: u64 = 904_538;
const GROUP_BY: u64 = 904_710;
const PIVOT_OWNER: u64 = 9_200_001;

fn rewrite_engine(edit: impl FnOnce(&mut Archive) -> TestResult) -> TestResult<Fixture> {
    let input = fixture("iwork/numbers/basic.numbers")?;
    let catalog = Catalog::from_bytes(&input.bytes, Default::default())?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == ENGINE)
        .ok_or("basic.numbers has no calculation engine")?;
    let stream = SnappyStream::decompress(entry.data(), Default::default())?;
    let mut archive = Archive::parse(stream.as_bytes(), Default::default())?;
    edit(&mut archive)?;
    let compressed =
        SnappyStream::compress(&archive.to_bytes(Default::default())?, Default::default())?;
    Ok(Fixture {
        path: "iwork/numbers/basic.numbers with table-owned pivot or category state",
        bytes: catalog
            .reassemble_to_bytes(&[EntryEdit::new(ENGINE, &compressed)], Default::default())?,
    })
}

/// The first table's category: its native group-by enabled.
fn active_category_fixture() -> TestResult<Fixture> {
    rewrite_engine(|archive| {
        let object = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(GROUP_BY))
            .ok_or("no native group-by")?;
        let mut group = tst::GroupByArchive::decode(object.messages[0].data.as_slice())?;
        group.is_enabled = true;
        object.messages[0].data = group.encode_to_vec();
        Ok(())
    })
}

/// The first table as a pivot table: its model names a pivot owner.
fn pivot_owner_fixture() -> TestResult<Fixture> {
    rewrite_engine(|archive| {
        let model = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(MODEL))
            .ok_or("no table model")?;
        let index = model
            .messages
            .iter()
            .position(|message| message.type_ == 6_001)
            .ok_or("no table-model message")?;
        let reference = tsp::Reference {
            identifier: PIVOT_OWNER,
            ..Default::default()
        }
        .encode_to_vec();
        // Field 85 (wire type 2), in field-number order as Numbers writes it.
        let mut pivot = vec![0xaa, 0x05, u8::try_from(reference.len())?];
        pivot.extend_from_slice(&reference);
        let source = model.messages[index].data.clone();
        let view = litchi_iwa_common::wire::WireView::parse(
            &source,
            litchi_iwa_common::WireLimits::DEFAULT,
        )?;
        let mut payload = Vec::with_capacity(source.len() + pivot.len());
        let mut inserted = false;
        for field in view.fields() {
            if !inserted && field.number() > 85 {
                payload.extend_from_slice(&pivot);
                inserted = true;
            }
            payload.extend_from_slice(field.raw());
        }
        if !inserted {
            payload.extend_from_slice(&pivot);
        }
        model.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 6_001,
                data: payload,
            },
            Default::default(),
        )?;
        let info = &mut model.archive_info.message_infos[index];
        info.object_references.push(PIVOT_OWNER);
        let mut field = FieldInfo::new(
            FieldPath::new(vec![85], Default::default())?,
            Default::default(),
        )?;
        field.object_references.push(PIVOT_OWNER);
        info.field_infos.push(field);
        archive.objects.push(ArchiveObject::new(
            PIVOT_OWNER,
            vec![RawMessage {
                type_: 6_370,
                data: tst::PivotOwnerArchive {
                    pivot_owner_uid: Some(tsp::Uuid {
                        lower: 69,
                        upper: 70,
                    }),
                    ..Default::default()
                }
                .encode_to_vec(),
            }],
            Default::default(),
        )?);
        Ok(())
    })
}

/// The first table's deprecated inline category owner (model field 81):
/// its group-by enabled (`is_enabled`, field 6, 0 -> 1 in place).
fn deprecated_category_fixture() -> TestResult<Fixture> {
    use litchi_iwa_common::{WireLimits, wire::WireView};

    rewrite_engine(|archive| {
        let model = archive
            .objects
            .iter_mut()
            .find(|object| object.archive_info.identifier == Some(MODEL))
            .ok_or("no table model")?;
        let index = model
            .messages
            .iter()
            .position(|message| message.type_ == 6_001)
            .ok_or("no table-model message")?;
        let mut payload = model.messages[index].data.clone();
        let offset = {
            let base = payload.as_ptr() as usize;
            let view = WireView::parse(&payload, WireLimits::DEFAULT)?;
            let owner = view
                .fields()
                .find(|field| field.number() == 81)
                .ok_or("no deprecated category owner")?;
            let owner = WireView::parse(owner.payload(), WireLimits::DEFAULT)?;
            let group = owner
                .fields()
                .find(|field| field.number() == 2)
                .ok_or("no inline group-by")?;
            let group = WireView::parse(group.payload(), WireLimits::DEFAULT)?;
            let enabled = group
                .fields()
                .find(|field| field.number() == 6)
                .ok_or("no is_enabled")?;
            if enabled.payload() != [0] {
                return Err("inline group-by is already enabled".into());
            }
            enabled.payload().as_ptr() as usize - base
        };
        payload[offset] = 1;
        model.replace_message_preserving_header(
            index,
            RawMessage {
                type_: 6_001,
                data: payload,
            },
            Default::default(),
        )?;
        Ok(())
    })
}

/// C3 = B3 + 1 with a display cache.
#[inline(never)]
fn host_route(source: &[u8]) -> TestResult<Vec<u8>> {
    let mut editor = numbers::Draft::from_bytes(source)?;
    let table = crate::numbers_cases::host_first_table_model_id(&editor)?;
    editor.set_formula_with_cached_value(
        table,
        2,
        2,
        FormulaExpression::binary(
            FormulaBinaryOperator::Add,
            FormulaExpression::Cell(FormulaCellReference::relative(2, 1)),
            FormulaExpression::Number(1.0),
        ),
        FormulaCachedValue::number(43.0)?,
    )?;
    Ok(editor.to_bytes()?)
}

fn leaf_commit(source: &[u8]) -> TestResult<litchi_numbers::table::cells::Commit> {
    let package = litchi_numbers::Package::from_bytes(source)?;
    Ok(package
        .edit_table_cells(0usize, 0usize)?
        .set_formula_cached(
            At::new(2, 2),
            Expression::binary(
                BinaryOperator::Add,
                Expression::cell(CellReference::relative(2, 1)),
                Expression::number(1.0)?,
            )?,
            CachedValue::number(43.0)?,
        )?
        .commit()?)
}

fn leaf_route(source: &[u8]) -> TestResult<LeafRun> {
    let commit = leaf_commit(source)?;
    let restored = commit
        .package()
        .apply_table_cells(&commit.patch().inverse())?;
    Ok(LeafRun::with_inverse(
        leaf_bytes!(commit.package()),
        leaf_bytes!(restored.package()),
    ))
}

#[inline(never)]
fn owned_leaf_route(source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(leaf_bytes!(leaf_commit(source)?.package()))
}

fn cell_projection(bytes: &[u8]) -> TestResult<String> {
    let package = litchi_numbers::Package::from_bytes(bytes)?;
    Ok(format!(
        "{:?}",
        package.table_cell(0usize, 0usize, At::new(2, 2))?.storage()
    ))
}

fn run(name: &'static str, input: &Fixture) -> TestResult {
    Case::new(name, Format::Numbers, Tier::B, Evidence::Queued, input)
        .leaf_structural_metadata()
        .leaf_invalidates_root_previews()
        .host(host_route)
        .leaf(leaf_route)
        .read("C3 cell", cell_projection)
        .run()?;
    Ok(())
}

/// An active category (the table's category owner names an enabled
/// group-by): the host writes the cell, and so does the leaf.
#[test]
fn numbers_active_category_table_cell_parity() -> TestResult {
    run(
        "numbers_active_category_table_cell_parity",
        &active_category_fixture()?,
    )
}

/// An active deprecated inline category owner.
#[test]
fn numbers_deprecated_category_table_cell_parity() -> TestResult {
    run(
        "numbers_deprecated_category_table_cell_parity",
        &deprecated_category_fixture()?,
    )
}

/// A pivot table (the model names a pivot owner).
#[test]
fn numbers_pivot_owner_table_cell_parity() -> TestResult {
    run(
        "numbers_pivot_owner_table_cell_parity",
        &pivot_owner_fixture()?,
    )
}

/// A value write (no host Numbers value setter remains to compare
/// against) succeeds on the same tables and reads back.
#[test]
fn numbers_owned_table_value_write_is_admitted() -> TestResult {
    for input in [
        active_category_fixture()?,
        deprecated_category_fixture()?,
        pivot_owner_fixture()?,
    ] {
        let package = litchi_numbers::Package::from_bytes(&input.bytes)?;
        let commit = package
            .edit_table_cells(0usize, 0usize)?
            .set(
                At::new(2, 2),
                litchi_numbers::table::cells::Input::number(5.0)?,
            )?
            .commit()?;
        let output = litchi_numbers::Package::from_bytes(&leaf_bytes!(commit.package()))?;
        let state = output.table_cell(0usize, 0usize, At::new(2, 2))?;
        assert!(matches!(
            state.storage(),
            litchi_numbers::table::cells::Storage::Stored(litchi_numbers::cell::Value::Number(value))
                if value.get() == 5.0
        ));
    }
    Ok(())
}

/// Callgrind probe (record 0424a): the cached formula on the active-category
/// table, host against leaf, ABBA (`*owned_host_route*` /
/// `*owned_leaf_route*`).
#[test]
#[ignore = "callgrind cost probe; run explicitly"]
fn numbers_owned_table_cell_cost_abba() -> TestResult {
    let input = active_category_fixture()?;
    for route in [0_u8, 1, 1, 0] {
        if route == 0 {
            std::hint::black_box(owned_host_route(&input.bytes)?);
        } else {
            std::hint::black_box(owned_leaf_route(&input.bytes)?);
        }
    }
    Ok(())
}

#[inline(never)]
fn owned_host_route(source: &[u8]) -> TestResult<Vec<u8>> {
    host_route(source)
}
