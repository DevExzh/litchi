//! Focused regression tests for the MS-XLDM outer-storage facade.

use super::codec::{crc32, utf16le};
use super::model::{BOM, CRC_SIZE};
use super::{
    FileGroupClass, FileKind, GeneratedNameKind, XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE,
    classify_generated_path, inspect, write,
};

#[cfg(test)]
pub(crate) fn test_xldm_bytes() -> Vec<u8> {
    let payload = b"compressed-looking model metadata";
    let log = test_backup_log("Model.1.db.xml", payload.len() as i32, 100002);
    build_test_storage(&[
        ("Partitions", partitions_xml().as_bytes()),
        ("Model.1.db.xml", payload),
        ("BackupLog", log.as_bytes()),
    ])
}

#[cfg(test)]
fn partitions_xml() -> String {
    "<Partitions><Partition><ObjectPath></ObjectPath><Name></Name><DataSize>0</DataSize><Location></Location><DataSourceID></DataSourceID><ConnectionString></ConnectionString></Partition></Partitions>".into()
}

#[cfg(test)]
fn test_backup_log(path: &str, size: i32, class: i32) -> String {
    format!(
        "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>Model</ObjectId><Write>ReadWrite</Write><OlapInfo>false</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups><FileGroup><Class>{class}</Class><ID>Model</ID><Name>Model</Name><ObjectVersion>1</ObjectVersion><PersistLocation>1</PersistLocation><PersistLocationPath></PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>11111111-2222-3333-4444-555555555555</ObjectID><FileList><BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile></FileList></FileGroup></FileGroups></BackupLog>"
    )
}

#[cfg(test)]
fn build_test_storage(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut bytes = vec![0; XLDM_PAGE_SIZE];
    bytes.extend_from_slice(&BOM);
    let mut allocations = Vec::new();
    for (index, (_, payload)) in entries.iter().enumerate() {
        if index + 1 == entries.len() {
            bytes.extend_from_slice(&BOM);
        }
        let offset = bytes.len();
        bytes.extend_from_slice(payload);
        bytes.extend_from_slice(&crc32(payload).to_le_bytes());
        allocations.push((offset, payload.len() + CRC_SIZE));
    }
    let directory_offset = bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE;
    bytes.resize(directory_offset, 0);
    let mut directory = String::from("<VirtualDirectory>");
    for ((path, _), (offset, size)) in entries.iter().zip(&allocations) {
        directory.push_str(&format!("<BackupFile><Path>{path}</Path><Size>{size}</Size><m_cbOffsetHeader>{offset}</m_cbOffsetHeader><Delete>false</Delete><CreatedTimestamp>0</CreatedTimestamp><Access>0</Access><LastWriteTime>0</LastWriteTime></BackupFile>"));
    }
    directory.push_str("</VirtualDirectory>");
    let directory_bytes = utf16le(&directory);
    bytes.extend_from_slice(&directory_bytes);
    bytes.resize(bytes.len().div_ceil(XLDM_PAGE_SIZE) * XLDM_PAGE_SIZE, 0);
    let header = format!(
        "<BackupLog><BackupRestoreSyncVersion>140</BackupRestoreSyncVersion><Fault>false</Fault><faultcode>0</faultcode><ErrorCode>true</ErrorCode><EncryptionFlag>false</EncryptionFlag><EncryptionKey>0</EncryptionKey><ApplyCompression>true</ApplyCompression><m_cbOffsetHeader>{directory_offset}</m_cbOffsetHeader><DataSize>{}</DataSize><Files>{}</Files><ObjectID>01234567-89AB-CDEF-0123-456789ABCDEF</ObjectID><m_cbOffsetData>4096</m_cbOffsetData></BackupLog>",
        directory_bytes.len(),
        entries.len()
    );
    let mut page = Vec::new();
    page.extend_from_slice(&BOM);
    page.extend_from_slice(&utf16le(XLDM_STREAM_SIGNATURE));
    page.extend_from_slice(&utf16le(&header));
    page.resize(XLDM_PAGE_SIZE, 0);
    bytes[..XLDM_PAGE_SIZE].copy_from_slice(&page);
    bytes
}

#[cfg(test)]
#[allow(
    clippy::module_inception,
    reason = "the nested module scopes executable tests beside shared fixture helpers"
)]
mod tests {
    use super::*;

    #[test]
    fn inspects_typed_storage_and_writes_byte_exactly() {
        let bytes = test_xldm_bytes();
        let storage = inspect(&bytes).unwrap();
        assert_eq!(storage.header.backup_restore_sync_version, 140);
        assert_eq!(storage.partition_marker.partition_count, 1);
        assert_eq!(storage.files.len(), 3);
        assert_eq!(storage.files[0].kind, FileKind::Partitions);
        assert_eq!(storage.files[2].kind, FileKind::BackupLog);
        assert_eq!(
            storage.backup_log.file_groups[0].class,
            FileGroupClass::Database
        );
        assert_eq!(
            storage.backup_log.file_groups[0].files[0].generated.kind,
            GeneratedNameKind::DatabaseDefinition
        );
        assert_eq!(
            storage.file_payload(1).unwrap(),
            b"compressed-looking model metadata"
        );
        assert_eq!(write(&storage).unwrap(), bytes);
    }

    #[test]
    fn rejects_header_signature_version_padding_and_offsets() {
        let base = test_xldm_bytes();
        for mutation in [0usize, 4] {
            let mut bytes = base.clone();
            bytes[mutation] ^= 1;
            assert!(inspect(&bytes).is_err());
        }
        let mut bytes = base;
        bytes[XLDM_PAGE_SIZE - 1] = 1;
        assert!(inspect(&bytes).is_err());
    }

    #[test]
    fn rejects_crc_corruption_without_interpreting_payload() {
        let mut bytes = test_xldm_bytes();
        let storage = inspect(&bytes).unwrap();
        let offset = storage.files[1].offset.0 as usize;
        bytes[offset] ^= 1;
        assert!(inspect(&bytes).is_err());
    }

    #[test]
    fn rejects_overlap_gap_and_nonzero_page_padding() {
        let mut bytes = test_xldm_bytes();
        let storage = inspect(&bytes).unwrap();
        let padding = storage.files.last().unwrap().offset.0 as usize
            + storage.files.last().unwrap().stored_size.0 as usize;
        bytes[padding] = 1;
        assert!(inspect(&bytes).is_err());
    }

    #[test]
    fn rejects_traversal_duplicate_and_ancestor_cycles() {
        for paths in [
            ["Partitions", "../escape", "BackupLog"],
            ["Partitions", "A/B", "a/b"],
            ["Partitions", "A", "A/B"],
        ] {
            let entries: Vec<_> = paths
                .iter()
                .enumerate()
                .map(|(index, path)| {
                    (
                        *path,
                        if index == 0 {
                            partitions_xml().as_bytes().to_vec()
                        } else {
                            vec![index as u8]
                        },
                    )
                })
                .collect();
            let refs: Vec<_> = entries
                .iter()
                .map(|(path, data)| (*path, data.as_slice()))
                .collect();
            assert!(inspect(&build_test_storage(&refs)).is_err());
        }
    }

    #[test]
    fn rejects_directory_count_and_partition_shape_mismatches() {
        let bad_partition = b"<Partitions><Wrong/></Partitions>";
        assert!(
            inspect(&build_test_storage(&[
                ("Partitions", bad_partition),
                ("Data", b"x"),
                ("BackupLog", b"log")
            ]))
            .is_err()
        );
    }

    #[test]
    fn classifies_every_generated_filename_family() {
        let cases = [
            ("Model.1.db.xml", GeneratedNameKind::DatabaseDefinition),
            (
                "Model.1.db/View.2.dsv.xml",
                GeneratedNameKind::DataSourceViewDefinition,
            ),
            (
                "Model.1.db/Cube.2.cub.xml",
                GeneratedNameKind::CubeDefinition,
            ),
            (
                "Model.1.db/Cube.0.cub/info.3.xml",
                GeneratedNameKind::CubeInformation,
            ),
            (
                "Model.1.db/Cube.0.cub/MdxScript.0.scr.xml",
                GeneratedNameKind::MdxScriptMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det.xml",
                GeneratedNameKind::MeasureGroupMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det/Table.3.prt.xml",
                GeneratedNameKind::PartitionMetadata,
            ),
            (
                "Model.1.db/Cube.0.cub/Table.2.det/Table.4.prt/info.5.xml",
                GeneratedNameKind::PartitionInformation,
            ),
            (
                "Model.1.db/Table.0.dim/Table.1.tbl.xml",
                GeneratedNameKind::TableMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/R$Table$Rel.1.tbl.xml",
                GeneratedNameKind::TableRelationshipMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/H$Table$Col.1.tbl.xml",
                GeneratedNameKind::ColumnHierarchyMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/U$Table$Hier.1.tbl.xml",
                GeneratedNameKind::UserHierarchyMetadata,
            ),
            (
                "Model.1.db/Table.0.dim/1.Table.Col.0.idf",
                GeneratedNameKind::ColumnData,
            ),
            (
                "Model.1.db/Table.0.dim/1.R$Table$Rel.INDEX.0.idf",
                GeneratedNameKind::TableRelationshipIndex,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.POS_TO_ID.0.idf",
                GeneratedNameKind::ColumnPositionToId,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.ID_TO_POS.0.idf",
                GeneratedNameKind::ColumnIdToPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.H$Table$Col.hidx",
                GeneratedNameKind::ColumnHashIndex,
            ),
            (
                "Model.1.db/Table.0.dim/1.Table.Col.dictionary",
                GeneratedNameKind::ColumnDictionary,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.CHILD_COUNT.0.idf",
                GeneratedNameKind::UserHierarchyChildCount,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.FIRST_CHILD_POS.0.idf",
                GeneratedNameKind::UserHierarchyFirstChildPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.PARENT_POS.0.idf",
                GeneratedNameKind::UserHierarchyParentPosition,
            ),
            (
                "Model.1.db/Table.0.dim/1.U$Table$Hier.MULTI_LEVEL_ID.0.idf",
                GeneratedNameKind::UserHierarchyMultilevelId,
            ),
        ];
        for (path, expected) in cases {
            assert_eq!(
                classify_generated_path(path).unwrap().kind,
                expected,
                "{path}"
            );
        }
    }

    #[test]
    fn rejects_backup_log_size_class_and_name_mismatches() {
        let payload = b"data";
        for (path, size, class) in [
            ("Model.1.db.xml", 3, 100002),
            ("Model.1.db.xml", 4, 100006),
            ("../escape", 4, 100002),
        ] {
            let log = test_backup_log(path, size, class);
            let bytes = build_test_storage(&[
                ("Partitions", partitions_xml().as_bytes()),
                ("Model.1.db.xml", payload),
                ("BackupLog", log.as_bytes()),
            ]);
            assert!(inspect(&bytes).is_err());
        }
    }

    #[test]
    fn rejects_malformed_backup_log_flags_and_enums() {
        for replacement in [
            ("<SvrEncryptPwdFlag>true", "<SvrEncryptPwdFlag>false"),
            ("<Write>ReadWrite", "<Write>Execute"),
            ("<Class>100002", "<Class>100004"),
        ] {
            let payload = b"data";
            let log =
                test_backup_log("Model.1.db.xml", 4, 100002).replace(replacement.0, replacement.1);
            let bytes = build_test_storage(&[
                ("Partitions", partitions_xml().as_bytes()),
                ("Model.1.db.xml", payload),
                ("BackupLog", log.as_bytes()),
            ]);
            assert!(inspect(&bytes).is_err());
        }
    }
}
