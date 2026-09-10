//! Focused Pages merge readback coverage.
//!
//! The native fixture exercises the package ownership proof and the shared
//! borrowed merge decoder together.  Corruption cases rewrite only the
//! selected model payload, then assert that the public Pages API refuses the
//! result without exposing native identifiers.

use std::error::Error as StdError;

use litchi_iwa_archive::package::{Catalog, EntryEdit};
use litchi_iwa_common::table::merge::Region;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{tsce, tst};
use litchi_pages::{BodyTableMergesError, BodyTableSelector, Limits, Package};
use prost::Message as _;

const NATIVE: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-merges-native.pages"
));

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut output = Vec::new();
    package.write_to(&mut output)?;
    Ok(output)
}

fn rewrite_first_table_model(
    source: &[u8],
    mutate: impl FnOnce(&mut tst::TableModelArchive),
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut replacement = None;
    let mut mutate = Some(mutate);
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
        let mut changed = false;
        for object in &mut archive.objects {
            for message_index in 0..object.messages.len() {
                let message_type = object.messages[message_index].type_;
                let Ok(mut model) =
                    tst::TableModelArchive::decode(object.messages[message_index].data.as_slice())
                else {
                    continue;
                };
                if model.merge_owner.is_none() {
                    continue;
                }
                (mutate
                    .take()
                    .expect("the selected native model is visited only once"))(
                    &mut model
                );
                object.replace_message_preserving_header(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data: model.encode_to_vec(),
                    },
                )?;
                changed = true;
                break;
            }
            if changed {
                break;
            }
        }
        if changed {
            replacement = Some((
                entry.name().to_owned(),
                SnappyStream::compress(&archive.to_bytes()?)?.to_vec(),
            ));
            break;
        }
    }
    let Some((name, component)) = replacement else {
        return Err("native merge-owner model is missing".into());
    };
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(&name, component.as_slice())],
        Limits::default(),
    )?)
}

fn first_formula_mut(model: &mut tst::TableModelArchive) -> TestResult<&mut tsce::FormulaArchive> {
    model
        .merge_owner
        .as_mut()
        .and_then(|owner| owner.formula_store.as_mut())
        .and_then(|store| store.formulas.first_mut())
        .map(|pair| &mut pair.formula)
        .ok_or_else(|| "native merge formula is missing".into())
}

fn assert_invalid(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    assert_eq!(
        package.body_table_merges(BodyTableSelector::index(0)),
        Err(BodyTableMergesError::InvalidSource)
    );
    Ok(())
}

#[test]
fn native_merge_reads_by_name_and_position_without_rewriting_source() -> TestResult {
    let expected = vec![Region::new(3, 2, 1, 2)?];
    let package = Package::from_bytes(NATIVE)?;

    assert_eq!(package.body_table_merges("Table 1")?, expected);
    assert_eq!(
        package.body_table_merges(BodyTableSelector::index(0))?,
        expected
    );
    assert!(
        package
            .body_table_merges(BodyTableSelector::index(1))?
            .is_empty()
    );
    assert_eq!(
        package.body_table_merges(BodyTableSelector::name("missing")),
        Err(BodyTableMergesError::TableNotFound)
    );
    assert_eq!(exact_bytes(&package)?, NATIVE);
    Ok(())
}

#[test]
fn malformed_merge_formula_is_refused_before_publication() -> TestResult {
    let malformed = rewrite_first_table_model(NATIVE, |model| {
        let formula = first_formula_mut(model).expect("native formula exists");
        let function = formula
            .ast_node_array
            .ast_node
            .iter_mut()
            .find(|node| {
                node.ast_node_type == tsce::ast_node_array_archive::AstNodeType::FunctionNode as i32
            })
            .expect("native merge formula has a function node");
        function.ast_function_node_index = Some(167);
    })?;
    assert_invalid(&malformed)
}

#[test]
fn foreign_merge_formula_is_refused_before_publication() -> TestResult {
    let foreign = rewrite_first_table_model(NATIVE, |model| {
        let formula = first_formula_mut(model).expect("native formula exists");
        let range = formula
            .ast_node_array
            .ast_node
            .iter_mut()
            .find(|node| {
                node.ast_node_type
                    == tsce::ast_node_array_archive::AstNodeType::ColonTractNode as i32
            })
            .expect("native merge formula has a range node");
        let table_id = &mut range
            .ast_cross_table_reference_extra_info
            .as_mut()
            .expect("native merge formula has a table identity")
            .table_id;
        table_id.uuid_bytes = None;
        table_id.uuid_w0 = Some(1);
        table_id.uuid_w1 = Some(2);
        table_id.uuid_w2 = Some(3);
        table_id.uuid_w3 = Some(4);
    })?;
    assert_invalid(&foreign)
}

#[test]
fn overlapping_merge_regions_are_refused_without_partial_results() -> TestResult {
    let overlapping = rewrite_first_table_model(NATIVE, |model| {
        let owner = model
            .merge_owner
            .as_mut()
            .expect("native merge owner exists");
        let store = owner
            .formula_store
            .as_mut()
            .expect("native formula store exists");
        let pair = store
            .formulas
            .first()
            .cloned()
            .expect("native merge pair exists");
        let next = store.next_formula_index;
        store.next_formula_index = next.checked_add(1).expect("fixture index fits");
        let mut duplicate = pair;
        duplicate.formula_index = next;
        store.formulas.push(duplicate);
    })?;
    assert_invalid(&overlapping)
}

#[test]
fn empty_merge_owner_is_a_valid_rooted_empty_result() -> TestResult {
    let package = Package::from_bytes(NATIVE)?;
    assert!(
        package
            .body_table_merges(BodyTableSelector::index(1))?
            .is_empty()
    );
    Ok(())
}
