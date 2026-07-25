//! CFB-backed loading of complete inert MS-OVBA projects.

use super::{
    VbaDirectory, VbaError, VbaLimits, VbaModuleKind, check_limit, decompress_container, invalid,
};
use crate::OleFile;
use encoding_rs::Encoding;
use litchi_core::encoding::codepage_to_encoding;
use std::io::{Read, Seek};

const VBA_STORAGE_NAME: &str = "VBA";
const DIR_STREAM_NAME: &str = "dir";
const PROJECT_STREAM_NAME: &str = "PROJECT";

/// Raw and decoded text from an MS-OVBA stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaText {
    raw: Vec<u8>,
    text: String,
    had_decode_errors: bool,
}

impl VbaText {
    /// Original bytes after decompression, before character decoding.
    pub fn raw(&self) -> &[u8] {
        &self.raw
    }

    /// Text decoded with the project's declared code page.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Whether malformed byte sequences were replaced while decoding.
    pub fn had_decode_errors(&self) -> bool {
        self.had_decode_errors
    }

    fn decode(raw: Vec<u8>, encoding: &'static Encoding) -> Self {
        let (text, had_decode_errors) = encoding.decode_without_bom_handling(&raw);
        let text = text.into_owned();
        Self {
            raw,
            text,
            had_decode_errors,
        }
    }
}

/// One inert VBA module and its typed directory metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaModule {
    name: String,
    stream_name: String,
    text_offset: u32,
    kind: VbaModuleKind,
    read_only: bool,
    private: bool,
    source: VbaText,
}

impl VbaModule {
    /// VBA identifier for this module.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// CFB stream containing this module.
    pub fn stream_name(&self) -> &str {
        &self.stream_name
    }

    /// Byte offset at which compressed source begins in the module stream.
    pub fn text_offset(&self) -> u32 {
        self.text_offset
    }

    /// Broad module category from the `dir` stream.
    pub fn kind(&self) -> VbaModuleKind {
        self.kind
    }

    /// Whether this module is marked read-only.
    pub fn is_read_only(&self) -> bool {
        self.read_only
    }

    /// Whether this module is private to its project.
    pub fn is_private(&self) -> bool {
        self.private
    }

    /// Decompressed, inert module source.
    pub fn source(&self) -> &VbaText {
        &self.source
    }
}

/// A complete inert VBA project loaded from a CFB project-root storage.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct VbaProject {
    project_root_path: Vec<String>,
    code_page: u16,
    name: String,
    project_properties: VbaText,
    modules: Vec<VbaModule>,
}

impl VbaProject {
    /// Load an MS-OVBA project rooted at `project_root_path`.
    ///
    /// For an OOXML `vbaProject.bin` part, the project root is normally the
    /// CFB root and this path is empty. Legacy Excel normally passes
    /// `["_VBA_PROJECT_CUR"]`; other hosts use the storage discovered from
    /// their CFB directory.
    pub fn open<R: Read + Seek>(
        ole: &mut OleFile<R>,
        project_root_path: &[&str],
        limits: &VbaLimits,
    ) -> Result<Self, VbaError> {
        let mut dir_path = project_root_path.to_vec();
        dir_path.extend([VBA_STORAGE_NAME, DIR_STREAM_NAME]);
        let compressed_dir =
            read_limited_stream(ole, &dir_path, limits.max_compressed_stream_bytes)?;
        let directory = VbaDirectory::parse(&compressed_dir, limits)?;
        let encoding = codepage_to_encoding(u32::from(directory.code_page()))
            .ok_or(VbaError::UnsupportedCodePage(directory.code_page()))?;

        let mut project_path = project_root_path.to_vec();
        project_path.push(PROJECT_STREAM_NAME);
        let project_properties = VbaText::decode(
            read_limited_stream(ole, &project_path, limits.max_decompressed_stream_bytes)?,
            encoding,
        );

        let mut modules = Vec::with_capacity(directory.modules().len());
        let mut total_source_bytes = 0usize;
        for metadata in directory.modules() {
            let mut module_path = project_root_path.to_vec();
            module_path.extend([VBA_STORAGE_NAME, metadata.stream_name()]);
            let stream =
                read_limited_stream(ole, &module_path, limits.max_compressed_stream_bytes)?;
            let text_offset = usize::try_from(metadata.text_offset())
                .map_err(|_| invalid("module text offset does not fit usize"))?;
            let compressed_source = stream.get(text_offset..).ok_or_else(|| {
                invalid(format!(
                    "module {} text offset {} exceeds stream size {}",
                    metadata.name(),
                    text_offset,
                    stream.len()
                ))
            })?;
            let source_bytes = decompress_container(compressed_source, limits)?;
            total_source_bytes = total_source_bytes
                .checked_add(source_bytes.len())
                .ok_or_else(|| invalid("aggregate VBA source size overflow"))?;
            check_limit(
                "aggregate VBA module source bytes",
                total_source_bytes,
                limits.max_total_source_bytes,
            )?;
            modules.push(VbaModule {
                name: metadata.name().to_owned(),
                stream_name: metadata.stream_name().to_owned(),
                text_offset: metadata.text_offset(),
                kind: metadata.kind(),
                read_only: metadata.is_read_only(),
                private: metadata.is_private(),
                source: VbaText::decode(source_bytes, encoding),
            });
        }

        Ok(Self {
            project_root_path: project_root_path
                .iter()
                .map(|component| (*component).to_owned())
                .collect(),
            code_page: directory.code_page(),
            name: directory.project_name().to_owned(),
            project_properties,
            modules,
        })
    }

    /// CFB path of the MS-OVBA project root.
    pub fn project_root_path(&self) -> &[String] {
        &self.project_root_path
    }

    /// Project code page used to decode MBCS text.
    pub fn code_page(&self) -> u16 {
        self.code_page
    }

    /// VBA project identifier.
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Decoded text of the uncompressed `PROJECT` stream.
    pub fn project_properties(&self) -> &VbaText {
        &self.project_properties
    }

    /// Modules in `dir`-stream order.
    pub fn modules(&self) -> &[VbaModule] {
        &self.modules
    }
}

fn read_limited_stream<R: Read + Seek>(
    ole: &mut OleFile<R>,
    path: &[&str],
    maximum: usize,
) -> Result<Vec<u8>, VbaError> {
    let (parent, stream_name) = path
        .split_last()
        .map(|(last, parent)| (parent, *last))
        .ok_or_else(|| invalid("VBA stream path must not be empty"))?;
    let entry = ole
        .list_directory_entries(parent)?
        .into_iter()
        .find(|entry| entry.name.eq_ignore_ascii_case(stream_name))
        .ok_or(crate::OleError::StreamNotFound)?;
    let size =
        usize::try_from(entry.size).map_err(|_| invalid("VBA stream size does not fit usize"))?;
    check_limit("VBA CFB stream bytes", size, maximum)?;
    let stream = ole.open_stream(path)?;
    check_limit("VBA CFB stream bytes", stream.len(), maximum)?;
    Ok(stream)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::OleWriter;
    use std::io::Cursor;

    fn push_record(bytes: &mut Vec<u8>, id: u16, value: &[u8]) {
        bytes.extend_from_slice(&id.to_le_bytes());
        bytes.extend_from_slice(&(value.len() as u32).to_le_bytes());
        bytes.extend_from_slice(value);
    }

    fn push_pair(bytes: &mut Vec<u8>, id: u16, value: &str, reserved: u16) {
        push_record(bytes, id, value.as_bytes());
        bytes.extend_from_slice(&reserved.to_le_bytes());
        let unicode: Vec<u8> = value.encode_utf16().flat_map(u16::to_le_bytes).collect();
        bytes.extend_from_slice(&(unicode.len() as u32).to_le_bytes());
        bytes.extend_from_slice(&unicode);
    }

    fn push_project_information(bytes: &mut Vec<u8>, name: &str) {
        push_record(bytes, 0x0001, &1u32.to_le_bytes());
        push_record(bytes, 0x0002, &0x0409u32.to_le_bytes());
        push_record(bytes, 0x0014, &0x0409u32.to_le_bytes());
        push_record(bytes, 0x0003, &1252u16.to_le_bytes());
        push_record(bytes, 0x0004, name.as_bytes());
        push_pair(bytes, 0x0005, "", 0x0040);
        push_record(bytes, 0x0006, &[]);
        bytes.extend_from_slice(&0x003du16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        push_record(bytes, 0x0007, &0u32.to_le_bytes());
        push_record(bytes, 0x0008, &0u32.to_le_bytes());
        bytes.extend_from_slice(&0x0009u16.to_le_bytes());
        bytes.extend_from_slice(&4u32.to_le_bytes());
        bytes.extend_from_slice(&1u32.to_le_bytes());
        bytes.extend_from_slice(&0u16.to_le_bytes());
    }

    fn literal_container(data: &[u8]) -> Vec<u8> {
        let mut encoded = vec![0x01];
        for decompressed_chunk in data.chunks(3_000) {
            let mut chunk = Vec::with_capacity(decompressed_chunk.len() + 375);
            for literals in decompressed_chunk.chunks(8) {
                chunk.push(0);
                chunk.extend_from_slice(literals);
            }
            let header = 0xb000 | u16::try_from(chunk.len() - 1).unwrap();
            encoded.extend_from_slice(&header.to_le_bytes());
            encoded.extend_from_slice(&chunk);
        }
        encoded
    }

    fn sample_dir() -> Vec<u8> {
        let mut bytes = Vec::new();
        push_project_information(&mut bytes, "Sample");
        push_record(&mut bytes, 0x000f, &1u16.to_le_bytes());
        push_record(&mut bytes, 0x0013, &0xffffu16.to_le_bytes());
        push_record(&mut bytes, 0x0019, b"Module1");
        push_pair(&mut bytes, 0x001a, "Module1", 0x0032);
        push_pair(&mut bytes, 0x001c, "", 0x0048);
        push_record(&mut bytes, 0x0031, &3u32.to_le_bytes());
        push_record(&mut bytes, 0x001e, &0u32.to_le_bytes());
        push_record(&mut bytes, 0x002c, &0xffffu16.to_le_bytes());
        bytes.extend_from_slice(&0x0021u16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        bytes.extend_from_slice(&0x002bu16.to_le_bytes());
        bytes.extend_from_slice(&0u32.to_le_bytes());
        literal_container(&bytes)
    }

    #[test]
    fn opens_project_and_decompresses_inert_source() {
        let source = b"Attribute VB_Name = \"Module1\"\r\nSub Main()\r\nEnd Sub\r\n";
        let mut module_stream = vec![7, 8, 9];
        module_stream.extend_from_slice(&literal_container(source));

        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\nModule=Module1\r\n")
            .unwrap();
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .unwrap();
        writer
            .create_stream(&["VBA", "Module1"], &module_stream)
            .unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.set_position(0);

        let mut ole = OleFile::open(cursor).unwrap();
        let project = VbaProject::open(&mut ole, &[], &VbaLimits::default()).unwrap();
        assert_eq!(project.name(), "Sample");
        assert_eq!(project.code_page(), 1252);
        assert_eq!(project.modules().len(), 1);
        assert_eq!(project.modules()[0].source().raw(), source);
        assert_eq!(
            project.modules()[0].source().text(),
            std::str::from_utf8(source).unwrap()
        );
        assert!(!project.modules()[0].source().had_decode_errors());
        assert!(
            project
                .project_properties()
                .text()
                .contains("Module=Module1")
        );
    }

    #[test]
    fn rejects_module_offset_past_stream() {
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["PROJECT"], b"ID=\"Sample\"\r\n")
            .unwrap();
        writer
            .create_stream(&["VBA", "_VBA_PROJECT"], &[0; 8])
            .unwrap();
        writer
            .create_stream(&["VBA", "dir"], &sample_dir())
            .unwrap();
        writer.create_stream(&["VBA", "Module1"], &[1, 2]).unwrap();
        let mut cursor = Cursor::new(Vec::new());
        writer.write_to(&mut cursor).unwrap();
        cursor.set_position(0);
        let mut ole = OleFile::open(cursor).unwrap();
        assert!(VbaProject::open(&mut ole, &[], &VbaLimits::default()).is_err());
    }
}
