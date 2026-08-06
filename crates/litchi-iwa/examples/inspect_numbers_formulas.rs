use std::env;

use litchi_iwa::raw::package::IWorkPackage;
use litchi_iwa_protos::tsce::ast_node_array_archive::AstNodeType;
use litchi_iwa_protos::tst::TableDataList;
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_numbers_formulas <file.numbers>")?;
    let package = IWorkPackage::open(path)?;
    for archive_name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        let archive = package.archive(archive_name)?;
        for object in archive.objects {
            let object_id = object.archive_info.identifier.unwrap_or_default();
            for message in object.messages {
                if !matches!(message.type_, 6005 | 6201) {
                    continue;
                }
                let Ok(list) = TableDataList::decode(message.data.as_slice()) else {
                    continue;
                };
                for entry in list.entries {
                    let Some(formula) = entry.formula else {
                        continue;
                    };
                    println!(
                        "archive={archive_name} object={object_id} list_type={} key={} refs={} host=({:?},{:?}) flags={:?}",
                        list.list_type,
                        entry.key,
                        entry.refcount,
                        formula.host_column,
                        formula.host_row,
                        formula.translation_flags,
                    );
                    for (index, node) in formula.ast_node_array.ast_node.iter().enumerate() {
                        let kind = AstNodeType::try_from(node.ast_node_type)
                            .map(|kind| kind.as_str_name())
                            .unwrap_or("UNKNOWN");
                        println!(
                            " node[{index}]={kind} function={:?}/args={:?} number={:?} decimal={:?}/{:?} bool={:?} string={:?} coord={:?}/{:?} local_ref={:?} cross_ref={:?} cross_extra={:?} uid_ref={:?} category_ref={:?} category_levels={:?} sticky={:?} colon_tract={:?} tract={:?}",
                            node.ast_function_node_index,
                            node.ast_function_node_num_args,
                            node.ast_number_node_number,
                            node.ast_number_node_decimal_low,
                            node.ast_number_node_decimal_high,
                            node.ast_boolean_node_boolean,
                            node.ast_string_node_string,
                            node.ast_column,
                            node.ast_row,
                            node.ast_local_cell_reference_node_reference,
                            node.ast_cross_table_cell_reference_node_reference,
                            node.ast_cross_table_reference_extra_info,
                            node.ast_uid_coordinate,
                            node.ast_category_ref,
                            node.ast_category_levels,
                            node.ast_sticky_bits,
                            node.ast_colon_tract,
                            node.ast_tract_list,
                        );
                    }
                }
            }
        }
    }
    Ok(())
}
