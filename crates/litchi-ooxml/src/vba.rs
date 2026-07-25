//! Shared OOXML `vbaProject.bin` payload access.

use litchi_cfb::OleFile;
pub use litchi_cfb::ovba::{
    VbaDirectory, VbaError, VbaLimits, VbaModule, VbaModuleKind, VbaModuleMetadata, VbaProject,
    VbaText, compress_container, decompress_container,
};
use litchi_opc::{OpcPackage, PackURI};
use std::io::Cursor;

pub(crate) fn read_project_part(
    package: &OpcPackage,
    part_name: &PackURI,
    limits: &VbaLimits,
) -> Result<VbaProject, VbaError> {
    let part = package.get_part(part_name).map_err(|error| {
        VbaError::InvalidData(format!(
            "VBA Project part '{}' is unavailable: {error}",
            part_name.as_str()
        ))
    })?;
    let mut ole = OleFile::open(Cursor::new(part.blob()))?;
    VbaProject::open(&mut ole, &[], limits)
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_cfb::OleWriter;
    use litchi_opc::constants::content_type;
    use litchi_opc::part::BlobPart;

    fn push_record(bytes: &mut Vec<u8>, id: u16, value: &[u8]) {
        bytes.extend_from_slice(&id.to_le_bytes());
        bytes.extend_from_slice(&(value.len() as u32).to_le_bytes());
        bytes.extend_from_slice(value);
    }

    fn literal_container(data: &[u8]) -> Vec<u8> {
        let mut chunk = Vec::with_capacity(data.len() + data.len().div_ceil(8));
        for literals in data.chunks(8) {
            chunk.push(0);
            chunk.extend_from_slice(literals);
        }
        let mut encoded = vec![0x01];
        let header = 0xb000 | u16::try_from(chunk.len() - 1).unwrap();
        encoded.extend_from_slice(&header.to_le_bytes());
        encoded.extend_from_slice(&chunk);
        encoded
    }

    fn project_blob() -> Vec<u8> {
        let mut directory = Vec::new();
        push_record(&mut directory, 0x0001, &1u32.to_le_bytes());
        push_record(&mut directory, 0x0002, &0x0409u32.to_le_bytes());
        push_record(&mut directory, 0x0014, &0x0409u32.to_le_bytes());
        push_record(&mut directory, 0x0003, &1252u16.to_le_bytes());
        push_record(&mut directory, 0x0004, b"OoxmlProject");
        push_record(&mut directory, 0x0005, &[]);
        directory.extend_from_slice(&0x0040u16.to_le_bytes());
        directory.extend_from_slice(&0u32.to_le_bytes());
        push_record(&mut directory, 0x0006, &[]);
        directory.extend_from_slice(&0x003du16.to_le_bytes());
        directory.extend_from_slice(&0u32.to_le_bytes());
        push_record(&mut directory, 0x0007, &0u32.to_le_bytes());
        push_record(&mut directory, 0x0008, &0u32.to_le_bytes());
        directory.extend_from_slice(&0x0009u16.to_le_bytes());
        directory.extend_from_slice(&4u32.to_le_bytes());
        directory.extend_from_slice(&1u32.to_le_bytes());
        directory.extend_from_slice(&0u16.to_le_bytes());
        push_record(&mut directory, 0x000f, &0u16.to_le_bytes());
        push_record(&mut directory, 0x0013, &0xffffu16.to_le_bytes());

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"OoxmlProject\"\r\n")
            .unwrap();
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &literal_container(&directory))
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.into_inner()
    }

    #[test]
    fn reads_project_payload_directly_from_opc_part_without_copying_it() {
        let part_name = PackURI::new("/xl/vbaProject.bin").unwrap();
        let mut package = OpcPackage::new();
        package.add_part(Box::new(BlobPart::new(
            part_name.clone(),
            content_type::OFC_VBA_PROJECT.to_string(),
            project_blob(),
        )));

        let project = read_project_part(&package, &part_name, &VbaLimits::default()).unwrap();
        assert_eq!(project.name(), "OoxmlProject");
        assert_eq!(project.code_page(), 1252);
        assert!(project.modules().is_empty());
    }
}
