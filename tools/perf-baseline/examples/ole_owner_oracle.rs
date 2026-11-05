//! Build-free-from-the-runner preflight oracle for the W5 owner lane.
//!
//! This is intentionally a separate example from `ole_w5_probe`.  The latter
//! is the historical timing consumer and remains byte-for-byte immutable for
//! the baseline.  This example performs all correctness work before ABBA:
//! it records complete OLE stream inventories, checks every untouched sibling
//! and directory field, reads exact property values, and checks DOC/XLS edits
//! against explicit expected values.  It writes the source and edited bytes so
//! the Python post-run validator can hash the complete artifacts independently.

use std::error::Error as StdError;
use std::fs;
use std::hint::black_box;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_cfb::{OleFile, OleWriter};
use litchi_doc::tracked_revision::{RevisionKind, RevisionMetadata};
use litchi_doc::{Builder as DocBuilder, Package as DocPackage};
use litchi_ole_common::object::{Edit as ObjectEdit, Limits as ObjectLimits, Targets};
use litchi_ole_common::property_set::user_defined::{
    Edit as UserEdit, Hyperlink, Hyperlinks, Limits as UserLimits, LinkBase,
    Properties as UserProperties,
};
use litchi_ole_common::property_set::{CodePage, Section, USER_DEFINED_PROPERTIES_FMTID, Value};
use litchi_xls::Package as XlsPackage;
use litchi_xls::cell_values::{
    Reference as XlsReference, Selector as XlsSelector, Value as XlsValue,
};
use sha2::{Digest, Sha256};

#[cfg(feature = "current-owner-api")]
use litchi_ole_common::object::{OpenOptions, Package as ObjectPackage};

#[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
type OwnerSource = litchi_ole_common::object::Snapshot;
#[cfg(feature = "current-owner-api")]
type OwnerSource = litchi_ole_common::object::Package;

const TARGET_NAME: &str = "ole_w5_edit_target.bin";
const ORIGINAL: &[u8] = b"ole-w5-original";
const REPLACEMENT: &[u8] = b"ole-w5-replacement";
const SIBLINGS: usize = 64;
const DOC_REVISION_TEXT: &str = "ole-w5-tracked-revision ";
const DOC_REVISION_AUTHOR: &str = "W5Oracle";
const XLS_EDIT_VALUE: f64 = 4242.0;
const LINK_BASE_TEXT: &str = "https://example.invalid/base/";
const HYPERLINK_TARGET: &str = "https://example.invalid/target";
const HYPERLINK_LOCATION: &str = "bookmark";
const ORACLE_ARTIFACTS: &[&str] = &[
    "ole-source.bin",
    "ole-noop.bin",
    "ole-replacement.bin",
    "properties.tsv",
    "doc-source.bin",
    "doc-edit.bin",
    "doc-proof.tsv",
    "xls-source.bin",
    "xls-edit.bin",
    "xls-proof.tsv",
    "xls-opaque-source.bin",
    "xls-opaque-edit.bin",
    "doc-history-final.bin",
    "xls-history-final.bin",
    "history-proof.tsv",
    "ole-inventory.tsv",
    "oracle.tsv",
    "allocator.tsv",
];

struct Corpus {
    archive: Vec<u8>,
    limits: ObjectLimits,
    target_path: Vec<String>,
    user_section: Section,
    user_limits: UserLimits,
    doc_archive: Vec<u8>,
    xls_archive: Vec<u8>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct StreamRecord {
    name: String,
    entry_type: u8,
    sid: u32,
    sid_left: u32,
    sid_right: u32,
    sid_child: u32,
    clsid: String,
    state_bits: u32,
    creation_time: u64,
    modified_time: u64,
    size: u64,
    payload: Vec<u8>,
}

#[cfg(feature = "w5-counting-allocator")]
enum RetainedHistory {
    Doc(Vec<DocPackage>, Vec<litchi_doc::tracked_revision::Patch>),
    Xls(Vec<XlsPackage>, Vec<litchi_xls::Patch>),
}

#[cfg(feature = "w5-counting-allocator")]
mod counting {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicUsize, Ordering};

    static ALLOC_CALLS: AtomicUsize = AtomicUsize::new(0);
    static ALLOC_BYTES: AtomicUsize = AtomicUsize::new(0);
    static DEALLOC_CALLS: AtomicUsize = AtomicUsize::new(0);
    static DEALLOC_BYTES: AtomicUsize = AtomicUsize::new(0);
    static REALLOC_CALLS: AtomicUsize = AtomicUsize::new(0);
    static REALLOC_OLD_BYTES: AtomicUsize = AtomicUsize::new(0);
    static REALLOC_NEW_BYTES: AtomicUsize = AtomicUsize::new(0);
    static LIVE_CALLS: AtomicUsize = AtomicUsize::new(0);
    static LIVE_BYTES: AtomicUsize = AtomicUsize::new(0);
    static PEAK_CALLS: AtomicUsize = AtomicUsize::new(0);
    static PEAK_BYTES: AtomicUsize = AtomicUsize::new(0);

    pub struct Allocator;

    #[derive(Clone, Copy)]
    pub struct Snapshot {
        pub alloc_calls: usize,
        pub alloc_bytes: usize,
        pub dealloc_calls: usize,
        pub dealloc_bytes: usize,
        pub realloc_calls: usize,
        pub realloc_old_bytes: usize,
        pub realloc_new_bytes: usize,
        pub live_calls: usize,
        pub live_bytes: usize,
        pub peak_calls: usize,
        pub peak_bytes: usize,
    }

    #[global_allocator]
    static GLOBAL: Allocator = Allocator;

    #[inline]
    fn alloc(bytes: usize) {
        ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(bytes, Ordering::Relaxed);
        let calls = LIVE_CALLS.fetch_add(1, Ordering::Relaxed) + 1;
        let live = LIVE_BYTES.fetch_add(bytes, Ordering::Relaxed) + bytes;
        PEAK_CALLS.fetch_max(calls, Ordering::Relaxed);
        PEAK_BYTES.fetch_max(live, Ordering::Relaxed);
    }

    #[inline]
    fn dealloc(bytes: usize) {
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(bytes, Ordering::Relaxed);
        LIVE_CALLS.fetch_sub(1, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(bytes, Ordering::Relaxed);
    }

    #[inline]
    fn realloc(old_bytes: usize, new_bytes: usize) {
        REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        REALLOC_OLD_BYTES.fetch_add(old_bytes, Ordering::Relaxed);
        REALLOC_NEW_BYTES.fetch_add(new_bytes, Ordering::Relaxed);
        if new_bytes >= old_bytes {
            LIVE_BYTES.fetch_add(new_bytes - old_bytes, Ordering::Relaxed);
        } else {
            LIVE_BYTES.fetch_sub(old_bytes - new_bytes, Ordering::Relaxed);
        }
        PEAK_BYTES.fetch_max(LIVE_BYTES.load(Ordering::Relaxed), Ordering::Relaxed);
        PEAK_CALLS.fetch_max(LIVE_CALLS.load(Ordering::Relaxed), Ordering::Relaxed);
    }

    unsafe impl GlobalAlloc for Allocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() {
                alloc(layout.size());
            }
            pointer
        }

        unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc_zeroed(layout) };
            if !pointer.is_null() {
                alloc(layout.size());
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            unsafe { System.dealloc(pointer, layout) };
            dealloc(layout.size());
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            let next = unsafe { System.realloc(pointer, layout, new_size) };
            if !next.is_null() {
                realloc(layout.size(), new_size);
            }
            next
        }
    }

    pub fn begin() -> Snapshot {
        let live_calls = LIVE_CALLS.load(Ordering::Relaxed);
        let live_bytes = LIVE_BYTES.load(Ordering::Relaxed);
        PEAK_CALLS.store(live_calls, Ordering::Relaxed);
        PEAK_BYTES.store(live_bytes, Ordering::Relaxed);
        snapshot()
    }

    pub fn snapshot() -> Snapshot {
        Snapshot {
            alloc_calls: ALLOC_CALLS.load(Ordering::Relaxed),
            alloc_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Relaxed),
            dealloc_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
            realloc_calls: REALLOC_CALLS.load(Ordering::Relaxed),
            realloc_old_bytes: REALLOC_OLD_BYTES.load(Ordering::Relaxed),
            realloc_new_bytes: REALLOC_NEW_BYTES.load(Ordering::Relaxed),
            live_calls: LIVE_CALLS.load(Ordering::Relaxed),
            live_bytes: LIVE_BYTES.load(Ordering::Relaxed),
            peak_calls: PEAK_CALLS.load(Ordering::Relaxed),
            peak_bytes: PEAK_BYTES.load(Ordering::Relaxed),
        }
    }
}

fn main() -> Result<(), Box<dyn StdError>> {
    let output = arguments()?;
    fs::create_dir_all(&output)?;
    let corpus = build_corpus()?;

    let source = corpus.archive.clone();
    let noop = owner_commit(&corpus)?;
    let replacement = owner_replace(&corpus)?;
    check_ole_outputs(&source, &noop, &replacement)?;
    write(&output.join("ole-source.bin"), &source)?;
    write(&output.join("ole-noop.bin"), &noop)?;
    write(&output.join("ole-replacement.bin"), &replacement)?;

    let properties = check_properties(&corpus)?;
    write(&output.join("properties.tsv"), properties.as_bytes())?;

    let doc_source = DocPackage::from_bytes(&corpus.doc_archive)?;
    let (doc_output, doc_proof) = check_doc(&doc_source)?;
    write(&output.join("doc-source.bin"), &corpus.doc_archive)?;
    write(&output.join("doc-edit.bin"), &doc_output)?;
    write(&output.join("doc-proof.tsv"), doc_proof.as_bytes())?;

    let xls_source = XlsPackage::from_bytes(&corpus.xls_archive)?;
    let (xls_output, xls_proof, opaque_source, opaque_output) =
        check_xls(&xls_source, &corpus.xls_archive)?;
    write(&output.join("xls-source.bin"), &corpus.xls_archive)?;
    write(&output.join("xls-edit.bin"), &xls_output)?;
    write(&output.join("xls-proof.tsv"), xls_proof.as_bytes())?;
    write(&output.join("xls-opaque-source.bin"), &opaque_source)?;
    write(&output.join("xls-opaque-edit.bin"), &opaque_output)?;

    let (history_proof, doc_history, xls_history) = check_history(&corpus)?;
    write(&output.join("history-proof.tsv"), history_proof.as_bytes())?;
    write(&output.join("doc-history-final.bin"), &doc_history)?;
    write(&output.join("xls-history-final.bin"), &xls_history)?;

    let inventory_proof = inventory_proof(&source, &replacement)?;
    write(
        &output.join("ole-inventory.tsv"),
        inventory_proof.as_bytes(),
    )?;

    let record = format!(
        "schema\t1\nowner_streams\t{}\nowner_target\t{}\nowner_original_bytes\t{}\nowner_replacement_bytes\t{}\nproperties\tvalidated\ndoc\trevision\tInsertion\t0\t{}\t{}\nxls\tCells\t0\t0\t4242\n",
        SIBLINGS + 1,
        TARGET_NAME,
        ORIGINAL.len(),
        REPLACEMENT.len(),
        DOC_REVISION_TEXT.encode_utf16().count(),
        DOC_REVISION_AUTHOR,
    );
    write(&output.join("oracle.tsv"), record.as_bytes())?;

    #[cfg(feature = "w5-counting-allocator")]
    allocator_evidence(&output, &corpus)?;
    #[cfg(not(feature = "w5-counting-allocator"))]
    write(
        &output.join("allocator.tsv"),
        b"status\tnative-uninstrumented\n",
    )?;

    println!("ole_owner_oracle,{}", output.display());
    for name in ORACLE_ARTIFACTS {
        let bytes = fs::read(output.join(name))?;
        let digest = Sha256::digest(&bytes);
        println!(
            "oracle_output\t{name}\t{}\t{}",
            bytes.len(),
            digest
                .iter()
                .map(|byte| format!("{byte:02x}"))
                .collect::<String>(),
        );
    }
    Ok(())
}

fn arguments() -> Result<PathBuf, Box<dyn StdError>> {
    let mut output = None;
    let mut args = std::env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--output" => {
                output = Some(PathBuf::from(args.next().ok_or("--output needs a value")?))
            },
            "--help" | "-h" => {
                println!("usage: ole_owner_oracle --output DIRECTORY");
                std::process::exit(0);
            },
            other => return Err(format!("unknown argument {other:?}").into()),
        }
    }
    output.ok_or_else(|| "--output is required".into())
}

fn write(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn StdError>> {
    if path.exists() {
        return Err(format!("oracle output already exists: {}", path.display()).into());
    }
    fs::write(path, bytes)?;
    Ok(())
}

fn build_corpus() -> Result<Corpus, Box<dyn StdError>> {
    let mut writer = OleWriter::new();
    for index in 0..SIBLINGS {
        let name = format!("ole_w5_sibling_{index:03}");
        let payload = format!("opaque sibling payload {index:03}");
        writer.create_stream(&[name.as_str()], payload.as_bytes())?;
    }
    writer.create_stream(&[TARGET_NAME], ORIGINAL)?;
    let mut archive = Cursor::new(Vec::new());
    writer.write_to(&mut archive)?;
    let stream_count = u64::try_from(SIBLINGS + 1)?;
    let total_size = u64::try_from(archive.get_ref().len())?;
    let limits = ObjectLimits {
        max_objects: 1,
        max_storage_depth: 1,
        max_streams_per_object: usize::try_from(stream_count)?,
        max_streams: usize::try_from(stream_count)?,
        max_stream_size: total_size.max(64),
        max_object_size: total_size.saturating_add(4096),
        max_total_size: total_size.saturating_add(4096),
    };
    let mut user_section = Section::new(USER_DEFINED_PROPERTIES_FMTID);
    user_section.set_page(CodePage::Utf16Le);
    let user_limits = UserLimits::default();
    let link_base = LinkBase::new(LINK_BASE_TEXT)?;
    let hyperlinks = Hyperlinks::new(vec![Hyperlink::new(
        7,
        11,
        0,
        HYPERLINK_TARGET,
        HYPERLINK_LOCATION,
    )?]);
    {
        let mut edit = user_editor(&mut user_section, user_limits)?;
        edit.set_link_base(link_base)?;
        edit.set_hyperlinks(hyperlinks)?;
    }
    user_section.add(42, Value::Blob(vec![0x5a; 128]))?;
    Ok(Corpus {
        archive: archive.into_inner(),
        limits,
        target_path: vec![TARGET_NAME.to_owned()],
        user_section,
        user_limits,
        doc_archive: build_doc_archive()?,
        xls_archive: build_xls_archive()?,
    })
}

fn build_doc_archive() -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut builder = DocBuilder::new();
    for index in 0..256 {
        builder.add_paragraph(&format!("ole-w5 tracked paragraph {index:03}"))?;
    }
    Ok(builder.build()?.to_bytes()?)
}

fn build_xls_archive() -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut builder = litchi_xls::writer::Builder::new();
    let sheet = builder.add_worksheet("Cells")?;
    for row in 0..256u32 {
        builder.write_number(sheet, row, 0, f64::from(row))?;
    }
    let mut package = Cursor::new(Vec::new());
    builder.build()?.write_to(&mut package)?;
    let mut source = open_cfb(package.into_inner())?;
    let workbook = source.open_stream(&["Workbook"])?;
    let mut writer = OleWriter::new();
    writer.create_stream_owned(&["Workbook"], workbook)?;
    writer.create_stream_owned(&["ole_w5_opaque_payload.bin"], vec![0x5a; 96 * 1024])?;
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output)?;
    Ok(output.into_inner())
}

fn open_cfb(bytes: Vec<u8>) -> Result<OleFile<Cursor<Vec<u8>>>, litchi_cfb::OleError> {
    #[cfg(all(feature = "current-api", not(feature = "legacy-owner-api")))]
    {
        OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
    }
    #[cfg(not(all(feature = "current-api", not(feature = "legacy-owner-api"))))]
    {
        OleFile::open(Cursor::new(bytes))
    }
}

fn owner_source(corpus: &Corpus) -> Result<OwnerSource, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(ObjectEdit::open(corpus.archive.clone(), Targets::default(), corpus.limits)?.staged())
    }
    #[cfg(feature = "current-owner-api")]
    {
        Ok(ObjectPackage::open_with_options(
            corpus.archive.clone(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(corpus.limits),
        )?)
    }
}

fn owner_edit(corpus: &Corpus) -> Result<ObjectEdit, Box<dyn StdError>> {
    let source = owner_source(corpus)?;
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(source.edit())
    }
    #[cfg(feature = "current-owner-api")]
    {
        Ok(source.edit()?)
    }
}

fn owner_commit(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn StdError>> {
    let package = owner_edit(corpus)?.commit()?.into_parts().0;
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(package.finish()?)
    }
    #[cfg(feature = "current-owner-api")]
    {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

fn owner_replace(corpus: &Corpus) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut edit = owner_edit(corpus)?;
    edit.put_stream(&corpus.target_path, REPLACEMENT.to_vec())?;
    let package = edit.commit()?.into_parts().0;
    #[cfg(all(feature = "legacy-owner-api", not(feature = "current-owner-api")))]
    {
        Ok(package.finish()?)
    }
    #[cfg(feature = "current-owner-api")]
    {
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

fn inventory(bytes: &[u8]) -> Result<Vec<StreamRecord>, Box<dyn StdError>> {
    let mut file = open_cfb(bytes.to_vec())?;
    let mut records = Vec::new();
    if let Some(entry) = file.directory_entry_by_sid(0).cloned() {
        records.push(StreamRecord {
            name: entry.name,
            entry_type: entry.entry_type,
            sid: entry.sid,
            sid_left: entry.sid_left,
            sid_right: entry.sid_right,
            sid_child: entry.sid_child,
            clsid: entry.clsid,
            state_bits: entry.state_bits,
            creation_time: entry.creation_time,
            modified_time: entry.modified_time,
            size: entry.size,
            payload: Vec::new(),
        });
    }
    file.visit_directory_entries(&[], SIBLINGS + 4, |file, sid| {
        let entry = file
            .directory_entry_by_sid(sid)
            .cloned()
            .ok_or_else(|| "directory visitor returned an unknown SID".to_string())?;
        let name = entry.name.clone();
        let payload = if entry.entry_type == 2 {
            file.open_stream(&[name.as_str()])?
        } else {
            Vec::new()
        };
        let record = StreamRecord {
            name: name.clone(),
            entry_type: entry.entry_type,
            sid: entry.sid,
            sid_left: entry.sid_left,
            sid_right: entry.sid_right,
            sid_child: entry.sid_child,
            clsid: entry.clsid.clone(),
            state_bits: entry.state_bits,
            creation_time: entry.creation_time,
            modified_time: entry.modified_time,
            size: entry.size,
            payload,
        };
        records.push(record);
        Ok::<(), Box<dyn StdError>>(())
    })?;
    records.sort_by(|left, right| left.name.cmp(&right.name));
    Ok(records)
}

fn metadata_equal(left: &StreamRecord, right: &StreamRecord) -> bool {
    left.name == right.name
        && left.entry_type == right.entry_type
        && left.sid == right.sid
        && left.sid_left == right.sid_left
        && left.sid_right == right.sid_right
        && left.sid_child == right.sid_child
        && left.clsid == right.clsid
        && left.state_bits == right.state_bits
        && left.creation_time == right.creation_time
        && left.modified_time == right.modified_time
        && left.size == right.size
}

fn check_ole_outputs(
    source: &[u8],
    noop: &[u8],
    replacement: &[u8],
) -> Result<(), Box<dyn StdError>> {
    if source != noop {
        return Err("owner no-op changed complete source bytes".into());
    }
    let source_records = inventory(source)?;
    let replacement_records = inventory(replacement)?;
    if source_records.len() != SIBLINGS + 2 || replacement_records.len() != source_records.len() {
        return Err("owner output stream inventory is incomplete".into());
    }
    for (before, after) in source_records.iter().zip(replacement_records.iter()) {
        if before.name == TARGET_NAME {
            if after.payload != REPLACEMENT || after.size != REPLACEMENT.len() as u64 {
                return Err(
                    "owner replacement target bytes differ from explicit expected bytes".into(),
                );
            }
        } else if !metadata_equal(before, after) || before.payload != after.payload {
            return Err(format!("untouched OLE sibling changed: {}", before.name).into());
        }
    }
    Ok(())
}

fn inventory_proof(source: &[u8], replacement: &[u8]) -> Result<String, Box<dyn StdError>> {
    let mut output = String::from(
        "kind\tname\tentry_type\tsid\tsid_left\tsid_right\tsid_child\tclsid\tstate_bits\tcreation_time\tmodified_time\tsize\tpayload_hex\n",
    );
    for (kind, records) in [
        ("source", inventory(source)?),
        ("replacement", inventory(replacement)?),
    ] {
        for record in records {
            output.push_str(&format!(
                "{kind}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\n",
                record.name,
                record.entry_type,
                record.sid,
                record.sid_left,
                record.sid_right,
                record.sid_child,
                record.clsid,
                record.state_bits,
                record.creation_time,
                record.modified_time,
                record.size,
                record
                    .payload
                    .iter()
                    .map(|byte| format!("{byte:02x}"))
                    .collect::<String>(),
            ));
        }
    }
    Ok(output)
}

fn user_properties<'a>(
    section: &'a Section,
    limits: UserLimits,
) -> Result<UserProperties<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-property-api")))]
    {
        Ok(UserProperties::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-property-api")]
    {
        Ok(UserProperties::new(section, limits)?)
    }
}

fn user_editor<'a>(
    section: &'a mut Section,
    limits: UserLimits,
) -> Result<UserEdit<'a>, Box<dyn StdError>> {
    #[cfg(all(feature = "legacy-property-api", not(feature = "current-property-api")))]
    {
        Ok(UserEdit::with_limits(section, limits)?)
    }
    #[cfg(feature = "current-property-api")]
    {
        Ok(UserEdit::new(section, limits)?)
    }
}

fn check_properties(corpus: &Corpus) -> Result<String, Box<dyn StdError>> {
    let properties = user_properties(&corpus.user_section, corpus.user_limits)?;
    let link_base = properties.link_base()?.ok_or("link base was omitted")?;
    if link_base.value() != LINK_BASE_TEXT {
        return Err("link base value differs from explicit expected value".into());
    }
    let hyperlinks = properties.hyperlinks()?.ok_or("hyperlinks were omitted")?;
    if hyperlinks.len() != 1 {
        return Err("hyperlink count differs from explicit expected value".into());
    }
    let link = &hyperlinks.links()[0];
    if link.app() != 7
        || link.office_art() != 11
        || link.info() != 0
        || link.target() != HYPERLINK_TARGET
        || link.location() != HYPERLINK_LOCATION
        || !link.hash_matches()
    {
        return Err("hyperlink property values differ from explicit expected values".into());
    }
    if corpus.user_section.property(42) != Some(&Value::Blob(vec![0x5a; 128])) {
        return Err("opaque user property was not retained".into());
    }
    let mut edited = corpus.user_section.clone();
    {
        let mut edit = user_editor(&mut edited, corpus.user_limits)?;
        edit.set_link_base(LinkBase::new(LINK_BASE_TEXT)?)?;
        edit.set_hyperlinks(Hyperlinks::new(vec![Hyperlink::new(
            7,
            11,
            0,
            HYPERLINK_TARGET,
            HYPERLINK_LOCATION,
        )?]))?;
    }
    if edited != corpus.user_section {
        return Err("property no-op changed unrelated section state".into());
    }
    Ok(format!(
        "link_base\t{}\nhyperlinks\t1\nhyperlink\t7\t11\t0\t{}\t{}\nopaque_property\t42\t128\n",
        link_base.value(),
        link.target(),
        link.location(),
    ))
}

fn check_doc(source: &DocPackage) -> Result<(Vec<u8>, String), Box<dyn StdError>> {
    if !source.tracked_revisions()?.is_empty() {
        return Err("DOC source unexpectedly contains tracked revisions".into());
    }
    let mut edit = source.edit_tracked_revisions()?;
    edit.add_text(
        0,
        DOC_REVISION_TEXT,
        RevisionKind::Insertion,
        RevisionMetadata::new(DOC_REVISION_AUTHOR),
    )?;
    let output = edit.commit()?.into_package();
    let revisions = output.tracked_revisions()?;
    let revision = revisions.first().ok_or("DOC edit omitted its revision")?;
    if revisions.len() != 1
        || revision.kind != RevisionKind::Insertion
        || revision.start_cp != 0
        || revision.end_cp != DOC_REVISION_TEXT.encode_utf16().count() as u32
        || revision.author != DOC_REVISION_AUTHOR
    {
        return Err("DOC revision metadata differs from explicit expected values".into());
    }
    let bytes = output.to_bytes()?;
    if bytes == source.to_bytes()? {
        return Err("DOC edit produced a source-identical output".into());
    }
    let source_paragraphs = source
        .document()
        .body_paragraphs()
        .map(|paragraph| paragraph.map(str::to_owned))
        .collect::<Result<Vec<_>, _>>()?;
    let output_paragraphs = output
        .document()
        .body_paragraphs()
        .map(|paragraph| paragraph.map(str::to_owned))
        .collect::<Result<Vec<_>, _>>()?;
    if source_paragraphs.len() != 256 || output_paragraphs.len() != 256 {
        return Err("DOC paragraph oracle did not retain all 256 paragraphs".into());
    }
    for index in 0..256 {
        let expected_source = format!("ole-w5 tracked paragraph {index:03}");
        let expected_output = if index == 0 {
            format!("{DOC_REVISION_TEXT}{expected_source}")
        } else {
            expected_source.clone()
        };
        if source_paragraphs[index] != expected_source
            || output_paragraphs[index] != expected_output
        {
            return Err(
                format!("DOC paragraph {index} differs from its explicit expected value").into(),
            );
        }
    }
    let mut proof = String::from("kind\tindex\ttext\n");
    for (kind, paragraphs) in [("source", source_paragraphs), ("edit", output_paragraphs)] {
        for (index, paragraph) in paragraphs.into_iter().enumerate() {
            proof.push_str(&format!("{kind}\t{index}\t{paragraph}\n"));
        }
    }
    Ok((bytes, proof))
}

fn check_xls(
    source: &XlsPackage,
    source_bytes: &[u8],
) -> Result<(Vec<u8>, String, Vec<u8>, Vec<u8>), Box<dyn StdError>> {
    let before = source
        .cell_values()?
        .worksheet(XlsSelector::Name("Cells"))?
        .ok_or("XLS source worksheet was omitted")?
        .number(XlsReference::new(0, 0)?)?
        .ok_or("XLS source cell was omitted")?;
    if before.value() != 0.0 {
        return Err("XLS source cell differs from explicit expected value".into());
    }
    let mut edit = source.edit()?;
    edit.set_value(
        XlsSelector::Name("Cells"),
        XlsReference::new(0, 0)?,
        XlsValue::Number(XLS_EDIT_VALUE),
    )?;
    let output = edit.commit()?.into_package();
    let after = output
        .cell_values()?
        .worksheet(XlsSelector::Name("Cells"))?
        .ok_or("XLS output worksheet was omitted")?
        .number(XlsReference::new(0, 0)?)?
        .ok_or("XLS output cell was omitted")?;
    if after.value() != XLS_EDIT_VALUE {
        return Err("XLS output cell differs from explicit expected value".into());
    }
    let bytes = output.to_bytes()?;
    if bytes == source.to_bytes()? {
        return Err("XLS edit produced a source-identical output".into());
    }
    let source_values = source.cell_values()?;
    let source_sheet = source_values
        .worksheet(XlsSelector::Name("Cells"))?
        .ok_or("XLS source worksheet was omitted")?;
    let output_values = output.cell_values()?;
    let output_sheet = output_values
        .worksheet(XlsSelector::Name("Cells"))?
        .ok_or("XLS output worksheet was omitted")?;
    let mut proof = String::from("kind\trow\tvalue_bits\n");
    for row in 0..256u32 {
        let source_number = source_sheet
            .number(XlsReference::new(row, 0)?)?
            .ok_or("XLS source cell inventory is incomplete")?;
        let output_number = output_sheet
            .number(XlsReference::new(row, 0)?)?
            .ok_or("XLS output cell inventory is incomplete")?;
        let expected_source = f64::from(row);
        let expected_output = if row == 0 {
            XLS_EDIT_VALUE
        } else {
            expected_source
        };
        if source_number.value() != expected_source || output_number.value() != expected_output {
            return Err(format!("XLS row {row} differs from its explicit expected value").into());
        }
        proof.push_str(&format!(
            "source\t{row}\t{}\n",
            source_number.value().to_bits()
        ));
        proof.push_str(&format!(
            "edit\t{row}\t{}\n",
            output_number.value().to_bits()
        ));
    }
    let mut source_cfb = open_cfb(source_bytes.to_vec())?;
    let mut output_cfb = open_cfb(bytes.clone())?;
    let opaque_source = source_cfb.open_stream(&["ole_w5_opaque_payload.bin"])?;
    let opaque_output = output_cfb.open_stream(&["ole_w5_opaque_payload.bin"])?;
    if opaque_source.len() != 96 * 1024 || opaque_output != opaque_source {
        return Err("XLS opaque 96 KiB stream was not retained exactly".into());
    }
    proof.push_str(&format!("opaque\tbytes\t{}\n", opaque_source.len()));
    Ok((bytes, proof, opaque_source, opaque_output))
}

fn check_history(corpus: &Corpus) -> Result<(String, Vec<u8>, Vec<u8>), Box<dyn StdError>> {
    let doc_source = DocPackage::from_bytes(&corpus.doc_archive)?;
    let (doc_packages, doc_patches) = retain_doc_history(&doc_source)?;
    let mut doc_forward = doc_source.clone();
    for patch in &doc_patches {
        doc_forward = doc_forward.apply_tracked_revisions(patch)?.into_package();
    }
    let doc_final = doc_forward.to_bytes()?;
    for patch in doc_patches.iter().rev() {
        doc_forward = doc_forward
            .apply_tracked_revisions(&patch.inverse())?
            .into_package();
    }
    if doc_forward.to_bytes()? != doc_source.to_bytes()? {
        return Err("DOC retained history inverse did not restore source bytes".into());
    }

    let xls_source = XlsPackage::from_bytes(&corpus.xls_archive)?;
    let (xls_packages, xls_patches) = retain_xls_history(&xls_source)?;
    let mut xls_forward = xls_source.clone();
    for patch in &xls_patches {
        xls_forward = xls_forward.apply(patch)?.into_package();
    }
    let xls_final = xls_forward.to_bytes()?;
    for patch in xls_patches.iter().rev() {
        xls_forward = xls_forward.apply(&patch.inverse())?.into_package();
    }
    if xls_forward.to_bytes()? != xls_source.to_bytes()? {
        return Err("XLS retained history inverse did not restore source bytes".into());
    }
    let proof = format!(
        "kind\tpackages\tpatches\tforward\tinverse\n\
doc\t{}\t{}\tvalidated\tvalidated\n\
xls\t{}\t{}\tvalidated\tvalidated\n",
        doc_packages.len(),
        doc_patches.len(),
        xls_packages.len(),
        xls_patches.len(),
    );
    Ok((proof, doc_final, xls_final))
}

#[cfg(feature = "w5-counting-allocator")]
fn allocator_evidence(output: &Path, corpus: &Corpus) -> Result<(), Box<dyn StdError>> {
    let routes = [
        "read",
        "edit",
        "xls_cells",
        "xls_cells_stage",
        "doc_history",
        "xls_history",
    ];
    let mut rows = Vec::new();
    for route in routes {
        let before = counting::begin();
        let (retained_packages, retained_patches, retained): (
            usize,
            usize,
            Option<RetainedHistory>,
        ) = match route {
            "read" => {
                let properties = user_properties(&corpus.user_section, corpus.user_limits)?;
                black_box((properties.link_base()?, properties.hyperlinks()?));
                (0, 0, None)
            },
            "edit" => {
                let mut section = corpus.user_section.clone();
                let mut edit = user_editor(&mut section, corpus.user_limits)?;
                edit.set_link_base(LinkBase::new(LINK_BASE_TEXT)?)?;
                edit.set_hyperlinks(Hyperlinks::new(vec![Hyperlink::new(
                    7,
                    11,
                    0,
                    HYPERLINK_TARGET,
                    HYPERLINK_LOCATION,
                )?]))?;
                black_box(section);
                (0, 0, None)
            },
            "xls_cells" | "xls_cells_stage" => {
                let package = XlsPackage::from_bytes(&corpus.xls_archive)?;
                let mut edit = package.edit()?;
                edit.set_value(
                    XlsSelector::Name("Cells"),
                    XlsReference::new(0, 0)?,
                    XlsValue::Number(XLS_EDIT_VALUE),
                )?;
                let commit = edit.commit()?;
                if route == "xls_cells" {
                    black_box(commit.into_package().to_bytes()?);
                } else {
                    black_box(commit);
                }
                (0, 0, None)
            },
            "doc_history" | "xls_history" => {
                // The retained-state routes intentionally keep the published
                // package bytes alive through the snapshot boundary, so the
                // resulting live/retained counters are observable rather than
                // a synthetic assertion.
                if route == "doc_history" {
                    let package = DocPackage::from_bytes(&corpus.doc_archive)?;
                    let (packages, patches) = retain_doc_history(&package)?;
                    let counts = (packages.len(), patches.len());
                    (
                        counts.0,
                        counts.1,
                        Some(RetainedHistory::Doc(packages, patches)),
                    )
                } else {
                    let package = XlsPackage::from_bytes(&corpus.xls_archive)?;
                    let (packages, patches) = retain_xls_history(&package)?;
                    let counts = (packages.len(), patches.len());
                    (
                        counts.0,
                        counts.1,
                        Some(RetainedHistory::Xls(packages, patches)),
                    )
                }
            },
            _ => unreachable!(),
        };
        let after = counting::snapshot();
        // Borrow the retained values after the snapshot so they remain live
        // through the counter capture instead of being consumed and dropped.
        black_box(&retained);
        rows.push((route, before, after, retained_packages, retained_patches));
    }
    let mut text = String::from(
        "route\talloc_calls\talloc_bytes\tdealloc_calls\tdealloc_bytes\trealloc_calls\trealloc_old_bytes\trealloc_new_bytes\tbase_live_calls\tbase_live_bytes\tlive_calls\tlive_bytes\tpeak_calls\tpeak_bytes\tretained_packages\tretained_patches\n",
    );
    for (route, before, after, retained_packages, retained_patches) in rows {
        text.push_str(&format!(
            "{route}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{}\t{retained_packages}\t{retained_patches}\n",
            after.alloc_calls.saturating_sub(before.alloc_calls),
            after.alloc_bytes.saturating_sub(before.alloc_bytes),
            after.dealloc_calls.saturating_sub(before.dealloc_calls),
            after.dealloc_bytes.saturating_sub(before.dealloc_bytes),
            after.realloc_calls.saturating_sub(before.realloc_calls),
            after.realloc_old_bytes.saturating_sub(before.realloc_old_bytes),
            after.realloc_new_bytes.saturating_sub(before.realloc_new_bytes),
            before.live_calls,
            before.live_bytes,
            after.live_calls,
            after.live_bytes,
            after.peak_calls,
            after.peak_bytes,
        ));
    }
    write(&output.join("allocator.tsv"), text.as_bytes())?;
    Ok(())
}

fn retain_doc_history(
    source: &DocPackage,
) -> Result<(Vec<DocPackage>, Vec<litchi_doc::tracked_revision::Patch>), Box<dyn StdError>> {
    let mut current = source.clone();
    let mut packages = Vec::with_capacity(129);
    let mut patches = Vec::with_capacity(128);
    packages.push(current.clone());
    for index in 0..128 {
        let mut edit = current.edit_tracked_revisions()?;
        if index == 0 {
            edit.add_text(
                0,
                DOC_REVISION_TEXT,
                RevisionKind::Insertion,
                RevisionMetadata::new(DOC_REVISION_AUTHOR),
            )?;
        } else {
            let revisions = edit.revisions()?;
            if revisions.len() != 1 {
                return Err("DOC retained history did not expose one revision".into());
            }
            edit.replace(
                index.min(revisions.len() - 1),
                RevisionMetadata::new(DOC_REVISION_AUTHOR),
            )?;
        }
        let commit = edit.commit()?;
        current = commit.package().clone();
        patches.push(commit.patch().clone());
        packages.push(current.clone());
    }
    Ok((packages, patches))
}

fn retain_xls_history(
    source: &XlsPackage,
) -> Result<(Vec<XlsPackage>, Vec<litchi_xls::Patch>), Box<dyn StdError>> {
    let mut current = source.clone();
    let mut packages = Vec::with_capacity(129);
    let mut patches = Vec::with_capacity(128);
    packages.push(current.clone());
    for row in 0..128u32 {
        let mut edit = current.edit()?;
        edit.set_value(
            XlsSelector::Name("Cells"),
            XlsReference::new(row, 0)?,
            XlsValue::Number(10_000.0 + f64::from(row)),
        )?;
        let commit = edit.commit()?;
        current = commit.package().clone();
        patches.push(commit.patch().clone());
        packages.push(current.clone());
    }
    Ok((packages, patches))
}
