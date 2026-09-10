//! Small canonical XLDM source fixture used by the XLSX Data Model tests.
#![cfg(test)]

use litchi_xldm::{XLDM_PAGE_SIZE, XLDM_STREAM_SIGNATURE};

const BOM: [u8; 2] = [0xFF, 0xFE];
const CRC_SIZE: usize = 4;

pub(crate) fn test_xldm_bytes() -> Vec<u8> {
    let payload = b"compressed-looking model metadata";
    let log = test_backup_log("Model.1.db.xml", payload.len() as i32, 100002);
    build_test_storage(&[
        ("Partitions", partitions_xml().as_bytes()),
        ("Model.1.db.xml", payload),
        ("BackupLog", log.as_bytes()),
    ])
}

fn partitions_xml() -> String {
    "<Partitions><Partition><ObjectPath></ObjectPath><Name></Name><DataSize>0</DataSize><Location></Location><DataSourceID></DataSourceID><ConnectionString></ConnectionString></Partition></Partitions>".into()
}

fn test_backup_log(path: &str, size: i32, class: i32) -> String {
    format!(
        "<BackupLog><BackupRestoreSyncVersion>1153</BackupRestoreSyncVersion><ServerRoot>C:\\inert</ServerRoot><SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML><ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>false</CompressionFlag><EncryptionFlag>false</EncryptionFlag><ObjectName>Model</ObjectName><ObjectId>Model</ObjectId><Write>ReadWrite</Write><OlapInfo>false</OlapInfo><Collations><Collation>Latin1_General</Collation></Collations><Languages><Language>1033</Language></Languages><FileGroups><FileGroup><Class>{class}</Class><ID>Model</ID><Name>Model</Name><ObjectVersion>1</ObjectVersion><PersistLocation>1</PersistLocation><PersistLocationPath></PersistLocationPath><StorageLocationPath></StorageLocationPath><ObjectID>11111111-2222-3333-4444-555555555555</ObjectID><FileList><BackupFile><Path>C:\\inert\\{path}</Path><StoragePath>{path}</StoragePath><LastWriteTime>0</LastWriteTime><Size>{size}</Size></BackupFile></FileList></FileGroup></FileGroups></BackupLog>"
    )
}

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

fn utf16le(value: &str) -> Vec<u8> {
    value.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xFFFF_FFFF;
    for byte in bytes {
        let mut index = ((crc >> 24) ^ u32::from(*byte)) & 0xFF;
        let mut table = index << 24;
        for _ in 0..8 {
            table = if table & 0x8000_0000 != 0 {
                (table << 1) ^ 0x04C1_1DB7
            } else {
                table << 1
            };
        }
        index = table;
        crc = (crc << 8) ^ index;
    }
    crc
}
