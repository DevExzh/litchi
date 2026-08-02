//! Inspect package media metadata and object-level data references.

use std::collections::HashMap;
use std::env;

use litchi_iwa::IWorkPackage;
use litchi_iwa::protobuf::tsp::{DataMetadataMap, PackageMetadata};
use prost::Message;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let path = env::args()
        .nth(1)
        .ok_or("usage: inspect_media_graph <iwork-file>")?;
    let package = IWorkPackage::open(path)?;
    let mut infos = HashMap::new();
    let mut metadata_map_id = None;
    for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let object_id = object.archive_info.identifier.unwrap_or_default();
            for (message, metadata) in object
                .messages
                .iter()
                .zip(&object.archive_info.message_infos)
            {
                if message.type_ == 11006
                    && let Ok(package_metadata) = PackageMetadata::decode(message.data.as_slice())
                {
                    println!(
                        "metadata_roundtrip_equal={} raw={} encoded={} data_metadata_map={:?}",
                        package_metadata.encode_to_vec() == message.data,
                        message.data.len(),
                        package_metadata.encoded_len(),
                        package_metadata
                            .data_metadata_map
                            .as_ref()
                            .map(|reference| reference.identifier),
                    );
                    metadata_map_id = package_metadata
                        .data_metadata_map
                        .as_ref()
                        .map(|reference| reference.identifier);
                    for data in package_metadata.datas {
                        infos.insert(data.identifier, data);
                    }
                }
                if !metadata.data_references.is_empty() {
                    println!(
                        "archive={name} object={object_id} type={} data_refs={:?}",
                        message.type_, metadata.data_references
                    );
                }
            }
        }
    }
    if let Some(metadata_map_id) = metadata_map_id {
        for name in package.entry_names().filter(|name| name.ends_with(".iwa")) {
            let archive = package.archive(name)?;
            if let Some(object) = archive.object(metadata_map_id) {
                for message in &object.messages {
                    let map = DataMetadataMap::decode(message.data.as_slice())?;
                    println!(
                        "data_metadata_map archive={name} object={metadata_map_id} type={} entries={:?}",
                        message.type_,
                        map.data_metadata_entries
                            .iter()
                            .map(|entry| entry.data_identifier)
                            .collect::<Vec<_>>()
                    );
                }
            }
        }
    }
    let mut infos = infos.into_values().collect::<Vec<_>>();
    infos.sort_by_key(|info| info.identifier);
    for info in infos {
        let name = info
            .file_name
            .as_deref()
            .filter(|name| !name.is_empty())
            .unwrap_or(&info.preferred_file_name);
        println!(
            "data={} name={name:?} preferred={:?} digest={} materialized={:?} remote_length={:?}",
            info.identifier,
            info.preferred_file_name,
            hex(&info.digest),
            info.materialized_length,
            info.remote_data_length,
        );
    }
    Ok(())
}

fn hex(bytes: &[u8]) -> String {
    const DIGITS: &[u8; 16] = b"0123456789abcdef";
    let mut output = String::with_capacity(bytes.len() * 2);
    for byte in bytes {
        output.push(DIGITS[(byte >> 4) as usize] as char);
        output.push(DIGITS[(byte & 0x0f) as usize] as char);
    }
    output
}
