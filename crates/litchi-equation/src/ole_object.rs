//! The compound file of an OLE equation object, as Office Open XML packages
//! store it (`word/embeddings/oleObjectN.bin`, `ppt/embeddings/…`,
//! `xl/embeddings/…`; ECMA-376 Part 1 §15.2.13) and binary files keep it as
//! an embedded storage: the `\u{1}CompObj` stream naming the class
//! ([MS-OLEDS] 2.3.8), the `\u{1}Ole` stream ([MS-OLEDS] 2.3.3), the
//! `\u{3}ObjInfo` stream and the `Equation Native` stream holding the
//! 28-byte `EQNOLEFILEHDR` and the MTEF data.

use std::io::Cursor;

use litchi_cfb::{OleFile, OleFileLimits, OleWriter};

use crate::{Error, MtefTarget, Result};

/// The largest `Equation Native` stream this module reads or writes.
pub const MAX_EQUATION_NATIVE_BYTES: u64 = 16 * 1024 * 1024;

/// The largest equation-object compound file this module opens.
pub const MAX_EQUATION_OBJECT_BYTES: u64 = 64 * 1024 * 1024;

/// The most storages and streams an equation object may hold. Reading and
/// replacing its `Equation Native` stream share this bound and
/// [`MAX_EQUATION_OBJECT_DEPTH`], so every object that reads can also be
/// rewritten.
pub const MAX_EQUATION_OBJECT_ENTRIES: usize = 4096;

/// The deepest storage nesting an equation object may have (a storage
/// directly below the root is at depth 1). Rebuilding a compound file costs
/// more than linear time in its nesting depth, so deeper objects are
/// refused, when read as when replaced.
pub const MAX_EQUATION_OBJECT_DEPTH: usize = 64;

const EQUATION_NATIVE: &str = "Equation Native";
const COMP_OBJ: &str = "\u{1}CompObj";
const OLE: &str = "\u{1}Ole";
const OBJ_INFO: &str = "\u{3}ObjInfo";
/// Cached presentations of an OLE object (`\u{2}OlePres000`, …).
const OLE_PRES_PREFIX: &str = "\u{2}OlePres";

/// The `\u{1}Ole` stream Word writes for an embedded MathType equation in
/// a DOCX package: version `0x02000001`, flags 8, no moniker.
const OLE_STREAM: [u8; 20] = [
    0x01, 0x00, 0x00, 0x02, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
];
/// The `\u{3}ObjInfo` stream Word writes beside it: clipboard format 3
/// (metafile picture).
const OBJ_INFO_STREAM: [u8; 6] = [0x00, 0x00, 0x03, 0x00, 0x01, 0x00];

impl MtefTarget {
    /// The clipboard format name the equation servers register.
    fn clipboard_format(&self) -> &'static str {
        if self.prog_id() == "Equation.3" {
            "DS Equation"
        } else {
            "MathType EF"
        }
    }

    /// The `CompObj` stream ([MS-OLEDS] 2.3.8) of a new object of this
    /// kind: the header, the class, and the user type, clipboard format and
    /// ProgID as ANSI strings, then the Unicode marker without Unicode
    /// strings.
    fn comp_obj(&self) -> Vec<u8> {
        let mut out = Vec::with_capacity(128);
        out.extend_from_slice(&[0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00]);
        out.extend_from_slice(&[0xFF; 4]);
        out.extend_from_slice(&self.clsid());
        for text in [self.user_type(), self.clipboard_format(), self.prog_id()] {
            let length = u32::try_from(text.len() + 1).unwrap_or(u32::MAX);
            out.extend_from_slice(&length.to_le_bytes());
            out.extend_from_slice(text.as_bytes());
            out.push(0);
        }
        out.extend_from_slice(&0x71B2_39F4u32.to_le_bytes());
        out.extend_from_slice(&[0; 12]);
        out
    }
}

/// The compound file of a new equation object of `target`'s kind holding
/// `equation_native` (a complete `Equation Native` stream, as
/// [`crate::latex_to_equation_native`] writes it): the root carries the
/// class identifier, and the streams are the ones Word writes for a
/// MathType object in a DOCX package.
///
/// # Errors
///
/// Returns [`Error::Object`] when the stream exceeds
/// [`MAX_EQUATION_NATIVE_BYTES`] or the compound file cannot be written.
pub fn equation_object(target: &MtefTarget, equation_native: &[u8]) -> Result<Vec<u8>> {
    check_native_size(equation_native.len())?;
    let mut writer = OleWriter::new();
    writer.set_root_clsid(target.clsid());
    writer
        .create_stream(&[COMP_OBJ], &target.comp_obj())
        .map_err(object_error)?;
    writer
        .create_stream(&[OLE], &OLE_STREAM)
        .map_err(object_error)?;
    writer
        .create_stream(&[OBJ_INFO], &OBJ_INFO_STREAM)
        .map_err(object_error)?;
    writer
        .create_stream(&[EQUATION_NATIVE], equation_native)
        .map_err(object_error)?;
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).map_err(object_error)?;
    Ok(out.into_inner())
}

/// The `Equation Native` stream of an equation object's compound file.
///
/// # Errors
///
/// Returns [`Error::Object`] when the data is not a compound file, is
/// larger than [`MAX_EQUATION_OBJECT_BYTES`], holds more than
/// [`MAX_EQUATION_OBJECT_ENTRIES`] storages and streams or nests them more
/// than [`MAX_EQUATION_OBJECT_DEPTH`] deep, has no top-level
/// `Equation Native` stream, or that stream is larger than
/// [`MAX_EQUATION_NATIVE_BYTES`].
pub fn equation_object_native(compound_file: &[u8]) -> Result<Vec<u8>> {
    let mut file = open(compound_file)?;
    collect_entries(&file)?;
    let size = file
        .list_directory_entries(&[])
        .map_err(object_error)?
        .into_iter()
        .find(|entry| entry.entry_type == STREAM && entry.name == EQUATION_NATIVE)
        .map(|entry| entry.size)
        .ok_or_else(|| Error::Object("the object has no Equation Native stream".to_owned()))?;
    check_native_size(usize::try_from(size).unwrap_or(usize::MAX))?;
    file.open_stream(&[EQUATION_NATIVE]).map_err(object_error)
}

/// `compound_file` (an equation object) with its `Equation Native` stream
/// replaced by `equation_native`. Cached presentations (`\u{2}OlePres*`)
/// are dropped, as they would show the old equation; every other stream,
/// storage, class identifier and state bit is copied.
///
/// # Errors
///
/// As [`equation_object_native`], and [`Error::Object`] when the compound
/// file cannot be written.
pub fn replace_equation_object_native(
    compound_file: &[u8],
    equation_native: &[u8],
) -> Result<Vec<u8>> {
    check_native_size(equation_native.len())?;
    let mut file = open(compound_file)?;
    let root = file
        .root_entry()
        .ok_or_else(|| Error::Object("the object has no root entry".to_owned()))?;
    let mut writer = OleWriter::new();
    if let Some(clsid) = parse_clsid(&root.clsid) {
        writer.set_root_clsid(clsid);
    }
    writer.set_root_state_bits(root.state_bits);
    let entries = collect_entries(&file)?;
    let mut wrote_native = false;
    for entry in entries {
        let refs = entry.path.iter().map(String::as_str).collect::<Vec<_>>();
        let top = entry.path.len() == 1;
        let name = entry.path.last().map_or("", String::as_str);
        if entry.storage {
            writer.create_storage(&refs).map_err(object_error)?;
            if let Some(clsid) = parse_clsid(&entry.clsid) {
                writer
                    .set_storage_clsid(&refs, clsid)
                    .map_err(object_error)?;
            }
            writer
                .set_storage_state_bits(&refs, entry.state_bits)
                .map_err(object_error)?;
        } else if top && name == EQUATION_NATIVE {
            writer
                .create_stream(&refs, equation_native)
                .map_err(object_error)?;
            wrote_native = true;
        } else if !(top && name.starts_with(OLE_PRES_PREFIX)) {
            let data = file.open_stream(&refs).map_err(object_error)?;
            writer.create_stream(&refs, &data).map_err(object_error)?;
        }
    }
    if !wrote_native {
        return Err(Error::Object(
            "the object has no Equation Native stream".to_owned(),
        ));
    }
    let mut out = Cursor::new(Vec::new());
    writer.write_to(&mut out).map_err(object_error)?;
    Ok(out.into_inner())
}

/// `STGTY_STORAGE` and `STGTY_STREAM` ([MS-CFB] 2.6.1).
const STORAGE: u8 = 1;
const STREAM: u8 = 2;

struct Entry {
    path: Vec<String>,
    storage: bool,
    clsid: String,
    state_bits: u32,
}

fn open(compound_file: &[u8]) -> Result<OleFile<Cursor<&[u8]>>> {
    if compound_file.len() as u64 > MAX_EQUATION_OBJECT_BYTES {
        return Err(Error::Object(format!(
            "the object's {} bytes exceed the {MAX_EQUATION_OBJECT_BYTES}-byte limit",
            compound_file.len()
        )));
    }
    let limits = OleFileLimits::new(MAX_EQUATION_OBJECT_BYTES).map_err(object_error)?;
    OleFile::open(Cursor::new(compound_file), limits).map_err(object_error)
}

/// Every storage and stream of the object, parents before children,
/// walked without recursion so that nesting depth costs no stack; refused
/// past [`MAX_EQUATION_OBJECT_ENTRIES`] entries or
/// [`MAX_EQUATION_OBJECT_DEPTH`] levels of storages.
fn collect_entries(file: &OleFile<Cursor<&[u8]>>) -> Result<Vec<Entry>> {
    let mut out: Vec<Entry> = Vec::new();
    // Storages whose children are still to be listed, by path.
    let mut pending: Vec<Vec<String>> = vec![Vec::new()];
    while let Some(parent) = pending.pop() {
        let refs = parent.iter().map(String::as_str).collect::<Vec<_>>();
        let children = file.list_directory_entries(&refs).map_err(object_error)?;
        // Children are pushed in reverse so storages are listed in order.
        let mut storages = Vec::new();
        for entry in children {
            let storage = match entry.entry_type {
                STORAGE => true,
                STREAM => false,
                _ => continue,
            };
            if out.len() >= MAX_EQUATION_OBJECT_ENTRIES {
                return Err(Error::Object(format!(
                    "the object holds more than {MAX_EQUATION_OBJECT_ENTRIES} storages and streams"
                )));
            }
            if storage && parent.len() >= MAX_EQUATION_OBJECT_DEPTH {
                return Err(Error::Object(format!(
                    "the object's storages nest more than {MAX_EQUATION_OBJECT_DEPTH} deep"
                )));
            }
            let mut path = parent.clone();
            path.push(entry.name.clone());
            if storage {
                storages.push(path.clone());
            }
            out.push(Entry {
                path,
                storage,
                clsid: entry.clsid.clone(),
                state_bits: entry.state_bits,
            });
        }
        pending.extend(storages.into_iter().rev());
    }
    Ok(out)
}

/// The 16 on-disk bytes of a CLSID written as `XXXXXXXX-XXXX-XXXX-XXXX-
/// XXXXXXXXXXXX` (the first three fields little-endian); `None` for an
/// empty or malformed one.
fn parse_clsid(text: &str) -> Option<[u8; 16]> {
    let hex: String = text.chars().filter(|c| *c != '-').collect();
    if hex.len() != 32 {
        return None;
    }
    let mut bytes = [0u8; 16];
    for (index, byte) in bytes.iter_mut().enumerate() {
        *byte = u8::from_str_radix(hex.get(index * 2..index * 2 + 2)?, 16).ok()?;
    }
    bytes[..4].reverse();
    bytes[4..6].reverse();
    bytes[6..8].reverse();
    Some(bytes)
}

fn check_native_size(len: usize) -> Result<()> {
    if len as u64 > MAX_EQUATION_NATIVE_BYTES {
        return Err(Error::Object(format!(
            "the Equation Native stream's {len} bytes exceed the \
             {MAX_EQUATION_NATIVE_BYTES}-byte limit"
        )));
    }
    Ok(())
}

fn object_error(error: litchi_cfb::OleError) -> Error {
    Error::Object(error.to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{Limits, latex_to_equation_native, mtef_to_latex};

    #[test]
    fn prog_ids_name_their_object_kind() {
        assert_eq!(
            MtefTarget::from_prog_id("Equation.DSMT4"),
            Some(MtefTarget::MathType)
        );
        assert_eq!(
            MtefTarget::from_prog_id("equation.dsmt6"),
            Some(MtefTarget::MathType)
        );
        assert_eq!(
            MtefTarget::from_prog_id("Equation.3"),
            Some(MtefTarget::EquationEditor3)
        );
        assert_eq!(MtefTarget::from_prog_id("Equation.2"), None);
        assert_eq!(MtefTarget::from_prog_id("Package"), None);
        assert_eq!(MtefTarget::from_prog_id(""), None);
    }

    #[test]
    fn comp_obj_matches_the_equation_editor_stream() {
        // The `CompObj` stream of an Equation Editor 3.0 object Word 2007
        // saved in a DOCX package (LibreOffice
        // `sw/qa/extras/ooxmlexport/data/mathtype.docx`).
        let mut expected = vec![
            0x01, 0x00, 0xFE, 0xFF, 0x03, 0x0A, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF,
        ];
        expected.extend_from_slice(&MtefTarget::EquationEditor3.clsid());
        expected.extend_from_slice(
            b"\x17\0\0\0Microsoft Equation 3.0\0\x0c\0\0\0DS Equation\0\x0b\0\0\0Equation.3\0",
        );
        expected.extend_from_slice(&[0xF4, 0x39, 0xB2, 0x71]);
        expected.extend_from_slice(&[0; 12]);
        assert_eq!(MtefTarget::EquationEditor3.comp_obj(), expected);
    }

    #[test]
    fn objects_are_written_read_and_replaced() {
        let limits = Limits::DEFAULT;
        let first = latex_to_equation_native("x^2", &MtefTarget::MathType, &limits).unwrap();
        let object = equation_object(&MtefTarget::MathType, first.bytes()).unwrap();
        assert_eq!(equation_object_native(&object).unwrap(), first.bytes());

        let second =
            latex_to_equation_native(r"\frac{a}{b}", &MtefTarget::MathType, &limits).unwrap();
        let replaced = replace_equation_object_native(&object, second.bytes()).unwrap();
        let native = equation_object_native(&replaced).unwrap();
        assert_eq!(native, second.bytes());
        assert_eq!(mtef_to_latex(&native, &limits).unwrap(), r"\frac{a}{b}");

        let mut file = open(&replaced).unwrap();
        let root = file.root_entry().unwrap().clsid.clone();
        assert_eq!(parse_clsid(&root), Some(MtefTarget::MathType.clsid()));
        assert_eq!(
            file.open_stream(&[COMP_OBJ]).unwrap(),
            MtefTarget::MathType.comp_obj()
        );
        assert_eq!(file.open_stream(&[OLE]).unwrap(), OLE_STREAM);
    }

    #[test]
    fn replacing_drops_cached_presentations_and_keeps_other_streams() {
        let mut writer = OleWriter::new();
        writer.set_root_clsid(MtefTarget::EquationEditor3.clsid());
        writer
            .create_stream(&[COMP_OBJ], &MtefTarget::EquationEditor3.comp_obj())
            .unwrap();
        writer
            .create_stream(&["\u{2}OlePres000"], &[1, 2, 3])
            .unwrap();
        writer.create_storage(&["Nested"]).unwrap();
        writer.create_stream(&["Nested", "Data"], &[7; 9]).unwrap();
        writer.create_stream(&[EQUATION_NATIVE], &[0; 40]).unwrap();
        let mut out = Cursor::new(Vec::new());
        writer.write_to(&mut out).unwrap();

        let replaced = replace_equation_object_native(&out.into_inner(), &[5; 50]).unwrap();
        let mut file = open(&replaced).unwrap();
        assert!(!file.exists(&["\u{2}OlePres000"]));
        assert_eq!(file.open_stream(&["Nested", "Data"]).unwrap(), [7; 9]);
        assert_eq!(file.open_stream(&[EQUATION_NATIVE]).unwrap(), [5; 50]);
        assert_eq!(
            parse_clsid(&file.root_entry().unwrap().clsid),
            Some(MtefTarget::EquationEditor3.clsid())
        );
    }

    /// An object whose storages nest `depth` deep, with a stream in the
    /// deepest one; returns it and that stream's path.
    fn nested(depth: usize) -> (Vec<u8>, Vec<String>) {
        let mut writer = OleWriter::new();
        writer.create_stream(&[EQUATION_NATIVE], &[0; 40]).unwrap();
        let names = (0..depth)
            .map(|level| format!("S{level}"))
            .collect::<Vec<_>>();
        for level in 1..=names.len() {
            let path = names[..level]
                .iter()
                .map(String::as_str)
                .collect::<Vec<_>>();
            writer.create_storage(&path).unwrap();
        }
        let mut leaf = names;
        leaf.push("Data".to_owned());
        let refs = leaf.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream(&refs, &[9; 3]).unwrap();
        let mut out = Cursor::new(Vec::new());
        writer.write_to(&mut out).unwrap();
        (out.into_inner(), leaf)
    }

    #[test]
    fn objects_nested_past_the_depth_limit_neither_read_nor_replace_quickly() {
        let (deep, _) = nested(MAX_EQUATION_OBJECT_DEPTH + 1);
        let started = std::time::Instant::now();
        assert!(matches!(
            equation_object_native(&deep),
            Err(Error::Object(_))
        ));
        assert!(matches!(
            replace_equation_object_native(&deep, &[1; 30]),
            Err(Error::Object(_))
        ));
        assert!(started.elapsed() < std::time::Duration::from_secs(5));
    }

    #[test]
    fn deeply_nested_objects_that_read_can_be_replaced() {
        let (deep, leaf) = nested(MAX_EQUATION_OBJECT_DEPTH);
        let leaf = leaf.iter().map(String::as_str).collect::<Vec<_>>();
        assert_eq!(equation_object_native(&deep).unwrap(), [0; 40]);
        let replaced = replace_equation_object_native(&deep, &[1; 30]).unwrap();
        let mut file = open(&replaced).unwrap();
        assert_eq!(file.open_stream(&leaf).unwrap(), [9; 3]);
        assert_eq!(equation_object_native(&replaced).unwrap(), [1; 30]);
    }

    #[test]
    fn objects_with_too_many_entries_neither_read_nor_replace() {
        let mut writer = OleWriter::new();
        writer.create_stream(&[EQUATION_NATIVE], &[0; 40]).unwrap();
        for index in 0..MAX_EQUATION_OBJECT_ENTRIES {
            writer.create_stream(&[&format!("s{index}")], &[]).unwrap();
        }
        let mut out = Cursor::new(Vec::new());
        writer.write_to(&mut out).unwrap();
        let crowded = out.into_inner();
        assert!(matches!(
            equation_object_native(&crowded),
            Err(Error::Object(_))
        ));
        assert!(matches!(
            replace_equation_object_native(&crowded, &[1; 30]),
            Err(Error::Object(_))
        ));
    }

    #[test]
    fn invalid_objects_are_refused() {
        assert!(matches!(
            equation_object_native(b"not a compound file"),
            Err(Error::Object(_))
        ));
        let mut writer = OleWriter::new();
        writer.create_stream(&["Other"], &[1]).unwrap();
        let mut out = Cursor::new(Vec::new());
        writer.write_to(&mut out).unwrap();
        let without_native = out.into_inner();
        assert!(matches!(
            equation_object_native(&without_native),
            Err(Error::Object(_))
        ));
        assert!(matches!(
            replace_equation_object_native(&without_native, &[0; 30]),
            Err(Error::Object(_))
        ));
        let huge = vec![0u8; usize::try_from(MAX_EQUATION_NATIVE_BYTES).unwrap() + 1];
        assert!(matches!(
            equation_object(&MtefTarget::MathType, &huge),
            Err(Error::Object(_))
        ));
    }
}
