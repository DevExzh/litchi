#![allow(unexpected_cfgs)]

//! Correctness and optional bounded phase timing for the ODF shared-triad lane.
//!
//! This probe is intentionally separate from the W8 109-route packet.  It
//! pairs the native owner and the `litchi_ops::File` route for the four ODF
//! object/text families changed by 0362usd.  The default `validate` mode is
//! untimed and proves source identity, exact no-ops, patch replay/inverse,
//! diagnostics and repeated publication.  `timed` is an execution aid for a
//! later quiet-host qualification: fixture I/O and owner opening stay outside
//! the measured closures, and the iteration count is bounded.  `allocator`
//! runs the same operation-scoped closures with the optional system allocator
//! observer enabled by the paired runner; its elapsed values are explicitly
//! non-claimable.

#[cfg(odf_shared_triad_allocator)]
mod allocation {
    use std::alloc::{GlobalAlloc, Layout, System};
    use std::sync::atomic::{AtomicBool, AtomicUsize, Ordering};

    pub(crate) const ENABLED: bool = true;

    static ACTIVE: AtomicBool = AtomicBool::new(false);
    static ALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
    static DEALLOCATIONS: AtomicUsize = AtomicUsize::new(0);
    static ALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
    static DEALLOCATED_BYTES: AtomicUsize = AtomicUsize::new(0);
    static REALLOCATIONS: AtomicUsize = AtomicUsize::new(0);

    pub struct CountingAllocator;

    // SAFETY: each operation delegates to the standard system allocator and
    // only updates independent atomic counters around the real call.
    unsafe impl GlobalAlloc for CountingAllocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            // SAFETY: the caller upholds GlobalAlloc's layout contract.
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() && ACTIVE.load(Ordering::Relaxed) {
                ALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                ALLOCATED_BYTES.fetch_add(layout.size(), Ordering::Relaxed);
            }
            pointer
        }

        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            // SAFETY: the caller upholds GlobalAlloc's pointer/layout contract.
            unsafe { System.dealloc(pointer, layout) };
            if ACTIVE.load(Ordering::Relaxed) {
                DEALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                DEALLOCATED_BYTES.fetch_add(layout.size(), Ordering::Relaxed);
            }
        }

        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
            // SAFETY: the caller upholds GlobalAlloc's pointer/layout contract.
            let replacement = unsafe { System.realloc(pointer, layout, new_size) };
            if !replacement.is_null() && ACTIVE.load(Ordering::Relaxed) {
                REALLOCATIONS.fetch_add(1, Ordering::Relaxed);
                if new_size >= layout.size() {
                    ALLOCATED_BYTES.fetch_add(new_size - layout.size(), Ordering::Relaxed);
                } else {
                    DEALLOCATED_BYTES.fetch_add(layout.size() - new_size, Ordering::Relaxed);
                }
            }
            replacement
        }
    }

    #[global_allocator]
    static GLOBAL: CountingAllocator = CountingAllocator;

    pub(crate) fn begin() {
        ACTIVE.store(false, Ordering::Relaxed);
        ALLOCATIONS.store(0, Ordering::Relaxed);
        DEALLOCATIONS.store(0, Ordering::Relaxed);
        ALLOCATED_BYTES.store(0, Ordering::Relaxed);
        DEALLOCATED_BYTES.store(0, Ordering::Relaxed);
        REALLOCATIONS.store(0, Ordering::Relaxed);
        ACTIVE.store(true, Ordering::Relaxed);
    }

    pub(crate) fn end() -> [usize; 5] {
        // Stop observing before loading the counters.  The caller may now
        // construct receipt JSON without charging that bookkeeping to the
        // operation that was measured.
        ACTIVE.store(false, Ordering::Relaxed);
        [
            ALLOCATIONS.load(Ordering::Relaxed),
            DEALLOCATIONS.load(Ordering::Relaxed),
            ALLOCATED_BYTES.load(Ordering::Relaxed),
            DEALLOCATED_BYTES.load(Ordering::Relaxed),
            REALLOCATIONS.load(Ordering::Relaxed),
        ]
    }
}

#[cfg(not(odf_shared_triad_allocator))]
mod allocation {
    pub(crate) const ENABLED: bool = false;

    pub(crate) fn begin() {}

    pub(crate) fn end() -> [usize; 5] {
        [0; 5]
    }
}

#[cfg(all(feature = "odc", feature = "odi", feature = "oth", feature = "odm"))]
mod probe {
    use std::fmt::Write as _;
    use std::hint::black_box;
    use std::io::{Cursor, Read};
    use std::time::Instant;

    use super::allocation;
    use litchi_core::Position;
    use litchi_odf_common::core::{PackageWriter, StreamingArchiveLimits, Structure};
    use litchi_ops::{File, TargetArg, open_bytes};
    use quick_xml::Reader;
    use quick_xml::events::Event;
    use serde_json::{Value, json};
    use sha2::{Digest as _, Sha256};
    use zip::{HasZipMetadata, ZipArchive};

    const MAX_ITERATIONS: usize = 128;
    const CHANGED_TEXT: &str = "shared-triad qualification";
    const ODC_PROPS: [(&str, &str); 1] = [("title", CHANGED_TEXT)];
    const ODI_PROPS: [(&str, &str); 4] = [
        ("x", "1cm"),
        ("y", "2cm"),
        ("width", "3cm"),
        ("height", "4cm"),
    ];
    const OTH_PROPS: [(&str, &str); 1] = [("text", CHANGED_TEXT)];
    const ODM_PROPS: [(&str, &str); 1] = [("title", CHANGED_TEXT)];

    fn digest(bytes: &[u8]) -> String {
        let mut output = String::with_capacity(64);
        for byte in Sha256::digest(bytes) {
            write!(output, "{byte:02x}").expect("digest formatting");
        }
        output
    }

    #[derive(Debug, PartialEq, Eq)]
    struct ArchiveMember {
        name: String,
        name_raw_sha256: String,
        content_sha256: String,
        content_bytes: usize,
        compression: String,
        compressed_bytes: u64,
        crc32: u32,
        version_needed: u16,
        flags: u16,
        compression_options: String,
        external_attributes: u32,
        is_utf8: bool,
        using_data_descriptor: bool,
        large_file: bool,
        central_extra_sha256: Option<String>,
        last_modified: Option<String>,
        extra_sha256: Option<String>,
        comment: String,
        unix_mode: Option<u32>,
        version_made_by: (u8, u8),
        encrypted: bool,
        content_text: Option<String>,
    }

    fn archive_members(bytes: &[u8]) -> (Vec<u8>, Vec<ArchiveMember>) {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("ODF ZIP archive");
        let archive_comment = archive.comment().to_vec();
        let mut members = Vec::with_capacity(archive.len());
        for index in 0..archive.len() {
            let mut member = archive.by_index(index).expect("ODF ZIP member");
            let name = member.name().expect("valid ZIP member name").into_owned();
            let name_raw_sha256 = digest(member.name_raw());
            let compression = format!("{:?}", member.compression());
            let compressed_bytes = member.compressed_size();
            let crc32 = member.crc32();
            let metadata = member.get_metadata();
            let version_needed = metadata.version_needed();
            let flags = metadata.flags.as_u16();
            let compression_options = format!("{:?}", member.options());
            let external_attributes = metadata.external_attributes;
            let is_utf8 = flags & 0x0800 != 0;
            let using_data_descriptor = metadata.is_using_data_descriptor();
            let large_file = metadata.large_file;
            // ZIP 9 exposes parsed extra fields. Hash the original central
            // record bytes so this preservation check remains byte-exact.
            let central_offset = usize::try_from(metadata.central_header_start)
                .expect("central header offset fits memory");
            let central = &bytes[central_offset..];
            let name_len = usize::from(u16::from_le_bytes(
                central[28..30].try_into().expect("central name length"),
            ));
            let extra_len = usize::from(u16::from_le_bytes(
                central[30..32].try_into().expect("central extra length"),
            ));
            let central_extra_sha256 = (extra_len != 0)
                .then(|| digest(&central[46 + name_len..46 + name_len + extra_len]));
            let last_modified = member.last_modified().map(|value| format!("{value:?}"));
            let extra_sha256 = member.extra_data().as_deref().map(digest);
            let comment = member.comment().to_owned();
            let unix_mode = member.unix_mode();
            let version_made_by = member.version_made_by();
            let encrypted = member.encrypted();
            let mut content = Vec::new();
            member
                .read_to_end(&mut content)
                .expect("ODF ZIP member content");
            let content_sha256 = digest(&content);
            let content_bytes = content.len();
            let content_text = if name.ends_with(".xml") {
                String::from_utf8(content).ok()
            } else {
                None
            };
            members.push(ArchiveMember {
                name,
                name_raw_sha256,
                content_sha256,
                content_bytes,
                compression,
                compressed_bytes,
                crc32,
                version_needed,
                flags,
                compression_options,
                external_attributes,
                is_utf8,
                using_data_descriptor,
                large_file,
                central_extra_sha256,
                last_modified,
                extra_sha256,
                comment,
                unix_mode,
                version_made_by,
                encrypted,
                content_text,
            });
        }
        (archive_comment, members)
    }

    fn archive_member_json(member: &ArchiveMember) -> Value {
        json!({
            "name": member.name,
            "name_raw_sha256": member.name_raw_sha256,
            "content_sha256": member.content_sha256,
            "content_bytes": member.content_bytes,
            "compression": member.compression,
            "compressed_bytes": member.compressed_bytes,
            "crc32": format!("{:08x}", member.crc32),
            "version_needed": member.version_needed,
            "flags": member.flags,
            "compression_options": member.compression_options,
            "external_attributes": member.external_attributes,
            "is_utf8": member.is_utf8,
            "using_data_descriptor": member.using_data_descriptor,
            "large_file": member.large_file,
            "central_extra_sha256": member.central_extra_sha256,
            "last_modified": member.last_modified,
            "extra_sha256": member.extra_sha256,
            "comment": member.comment,
            "unix_mode": member.unix_mode,
            "version_made_by": member.version_made_by,
            "encrypted": member.encrypted,
            "content_text": member.content_text,
        })
    }

    type ManifestAttributes = Vec<(String, String)>;

    #[derive(Debug, PartialEq, Eq, Clone)]
    struct ManifestSemantics {
        root_attributes: ManifestAttributes,
        ordered_entries: Vec<ManifestAttributes>,
        normalized_entries: Vec<ManifestAttributes>,
        other_events: Vec<String>,
    }

    fn manifest_attributes(
        element: &quick_xml::events::BytesStart<'_>,
        reader: &Reader<&[u8]>,
    ) -> ManifestAttributes {
        let mut attributes = element
            .attributes()
            .map(|attribute| {
                let attribute = attribute.expect("ODF manifest attribute");
                let key = String::from_utf8_lossy(attribute.key.as_ref()).into_owned();
                let value = attribute
                    .decoded_and_normalized_value(
                        quick_xml::XmlVersion::Implicit1_0,
                        reader.decoder(),
                    )
                    .expect("ODF manifest attribute value")
                    .into_owned();
                (key, value)
            })
            .collect::<Vec<_>>();
        attributes.sort();
        attributes
    }

    fn manifest_semantics(member: &ArchiveMember) -> ManifestSemantics {
        let xml = member
            .content_text
            .as_deref()
            .expect("ODF manifest XML is UTF-8");
        let mut reader = Reader::from_str(xml);
        let mut root_attributes = Vec::new();
        let mut ordered_entries = Vec::new();
        let mut other_events = Vec::new();
        let mut saw_root = false;
        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) | Ok(Event::Empty(element)) => {
                    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    let attributes = manifest_attributes(&element, &reader);
                    if !saw_root {
                        root_attributes = attributes;
                        saw_root = true;
                    } else if name.ends_with("file-entry") {
                        ordered_entries.push(attributes);
                    } else {
                        other_events.push(format!("start:{name}:{attributes:?}"));
                    }
                },
                Ok(Event::End(element)) => {
                    let name = String::from_utf8_lossy(element.name().as_ref()).into_owned();
                    if name != "manifest" && !name.ends_with(":manifest") {
                        other_events.push(format!("end:{name}"));
                    }
                },
                Ok(Event::Text(text)) => {
                    other_events.push(format!(
                        "text:{}",
                        text.decode().expect("ODF manifest text").into_owned()
                    ));
                },
                Ok(Event::CData(text)) => {
                    other_events.push(format!(
                        "cdata:{}",
                        text.decode().expect("ODF manifest CDATA").into_owned()
                    ));
                },
                Ok(Event::Comment(comment)) => {
                    other_events.push(format!(
                        "comment:{}",
                        comment.decode().expect("ODF manifest comment").into_owned()
                    ));
                },
                Ok(Event::Decl(declaration)) => {
                    other_events.push(format!("decl:{declaration:?}"));
                },
                Ok(Event::PI(instruction)) => {
                    other_events.push(format!("pi:{instruction:?}"));
                },
                Ok(Event::DocType(declaration)) => {
                    other_events.push(format!("doctype:{declaration:?}"));
                },
                Ok(Event::GeneralRef(reference)) => {
                    other_events.push(format!("general-ref:{reference:?}"));
                },
                Ok(Event::Eof) => break,
                Err(error) => panic!("ODF manifest XML: {error}"),
            }
        }
        assert!(saw_root, "ODF manifest has no root element");
        let mut normalized_entries = ordered_entries.clone();
        normalized_entries.sort();
        ManifestSemantics {
            root_attributes,
            ordered_entries,
            normalized_entries,
            other_events,
        }
    }

    fn manifest_semantics_json(semantics: &ManifestSemantics) -> Value {
        json!({
            "root_attributes": semantics.root_attributes,
            "ordered_entries": semantics.ordered_entries,
            "normalized_entries": semantics.normalized_entries,
            "other_events": semantics.other_events,
        })
    }

    fn manifest_paths(semantics: &ManifestSemantics) -> Vec<String> {
        semantics
            .ordered_entries
            .iter()
            .filter_map(|attributes| {
                attributes
                    .iter()
                    .find(|(key, _)| key == "manifest:full-path")
                    .map(|(_, value)| value.clone())
            })
            .collect()
    }

    fn payload_order(members: &[ArchiveMember]) -> Vec<String> {
        members
            .iter()
            .filter(|member| member.name != "mimetype" && member.name != "META-INF/manifest.xml")
            .map(|member| member.name.clone())
            .collect()
    }

    fn manifest_tracks_order(semantics: &ManifestSemantics, members: &[ArchiveMember]) -> bool {
        let manifest_payloads = manifest_paths(semantics)
            .into_iter()
            .filter(|path| path != "/")
            .collect::<Vec<_>>();
        manifest_payloads == payload_order(members)
    }

    fn archive_pair(leaf: &[u8], facade: &[u8]) -> Value {
        let (leaf_comment, leaf_members) = archive_members(leaf);
        let (facade_comment, facade_members) = archive_members(facade);
        assert_eq!(
            leaf_members.len(),
            facade_members.len(),
            "ODM archive member count differs"
        );
        assert_eq!(leaf_comment, facade_comment, "ODM archive comments differ");
        let leaf_comment_sha256 = digest(&leaf_comment);
        let facade_comment_sha256 = digest(&facade_comment);

        let mut member_rows = Vec::with_capacity(leaf_members.len());
        let mut member_order_exact = true;
        let mut non_meta_content_exact = true;
        let mut non_manifest_content_exact = true;
        let mut non_meta_metadata_exact = true;
        let mut non_manifest_metadata_exact = true;
        let mut non_manifest_record_exact = true;
        let mut meta_member_count = 0;
        let mut meta_metadata_exact = true;
        let mut meta_record_exact = true;
        let mut content_differences = Vec::new();
        let facade_by_name = facade_members
            .iter()
            .map(|member| (member.name.as_str(), member))
            .collect::<std::collections::HashMap<_, _>>();
        assert_eq!(
            facade_by_name.len(),
            facade_members.len(),
            "ODM facade archive contains duplicate member names"
        );
        let leaf_names = leaf_members
            .iter()
            .map(|member| member.name.as_str())
            .collect::<std::collections::HashSet<_>>();
        assert_eq!(
            leaf_names.len(),
            leaf_members.len(),
            "ODM leaf archive contains duplicate member names"
        );
        for (index, leaf_member) in leaf_members.iter().enumerate() {
            member_order_exact &= facade_members
                .get(index)
                .is_some_and(|member| member.name == leaf_member.name);
            let facade_member = facade_by_name
                .get(leaf_member.name.as_str())
                .copied()
                .expect("ODM archive member names");
            let content_equal = leaf_member.content_sha256 == facade_member.content_sha256;
            let metadata_equal = leaf_member.compression == facade_member.compression
                && leaf_member.name_raw_sha256 == facade_member.name_raw_sha256
                && leaf_member.compressed_bytes == facade_member.compressed_bytes
                && leaf_member.crc32 == facade_member.crc32
                && leaf_member.version_needed == facade_member.version_needed
                && leaf_member.flags == facade_member.flags
                && leaf_member.compression_options == facade_member.compression_options
                && leaf_member.external_attributes == facade_member.external_attributes
                && leaf_member.is_utf8 == facade_member.is_utf8
                && leaf_member.using_data_descriptor == facade_member.using_data_descriptor
                && leaf_member.large_file == facade_member.large_file
                && leaf_member.central_extra_sha256 == facade_member.central_extra_sha256
                && leaf_member.last_modified == facade_member.last_modified
                && leaf_member.extra_sha256 == facade_member.extra_sha256
                && leaf_member.comment == facade_member.comment
                && leaf_member.unix_mode == facade_member.unix_mode
                && leaf_member.version_made_by == facade_member.version_made_by
                && leaf_member.encrypted == facade_member.encrypted;
            let archive_record_equal = metadata_equal
                && leaf_member.content_bytes == facade_member.content_bytes
                && leaf_member.content_sha256 == facade_member.content_sha256;
            if leaf_member.name == "meta.xml" {
                meta_member_count += 1;
                meta_metadata_exact &= metadata_equal;
                meta_record_exact &= archive_record_equal;
            } else {
                non_meta_content_exact &= content_equal;
                if leaf_member.name != "META-INF/manifest.xml" {
                    non_manifest_content_exact &= content_equal;
                }
                non_meta_metadata_exact &= metadata_equal;
                if leaf_member.name != "META-INF/manifest.xml" {
                    non_manifest_metadata_exact &= metadata_equal;
                    non_manifest_record_exact &= archive_record_equal;
                }
            }
            if !content_equal {
                content_differences.push(leaf_member.name.clone());
            }
            member_rows.push(json!({
                "name": leaf_member.name,
                "facade_name": facade_member.name,
                "content_equal": content_equal,
                "stored_metadata_equal": metadata_equal,
                "archive_record_equal": archive_record_equal,
                "leaf": archive_member_json(leaf_member),
                "facade": archive_member_json(facade_member),
            }));
        }
        assert_eq!(
            meta_member_count, 1,
            "ODM archive must contain one meta.xml"
        );
        let manifest_content_difference =
            content_differences == [String::from("META-INF/manifest.xml")];
        let leaf_manifest = leaf_members
            .iter()
            .find(|member| member.name == "META-INF/manifest.xml")
            .expect("ODM leaf manifest");
        let facade_manifest = facade_members
            .iter()
            .find(|member| member.name == "META-INF/manifest.xml")
            .expect("ODM facade manifest");
        let leaf_manifest_semantics = manifest_semantics(leaf_manifest);
        let facade_manifest_semantics = manifest_semantics(facade_manifest);
        let manifest_semantics_equal = leaf_manifest_semantics.root_attributes
            == facade_manifest_semantics.root_attributes
            && leaf_manifest_semantics.normalized_entries
                == facade_manifest_semantics.normalized_entries
            && leaf_manifest_semantics.other_events == facade_manifest_semantics.other_events;
        let manifest_listing_tracks_order =
            manifest_paths(&leaf_manifest_semantics) == manifest_paths(&facade_manifest_semantics);
        let leaf_manifest_tracks_order =
            manifest_tracks_order(&leaf_manifest_semantics, &leaf_members);
        let facade_manifest_tracks_order =
            manifest_tracks_order(&facade_manifest_semantics, &facade_members);
        let archive_member_order_and_manifest_listing = !member_order_exact
            && manifest_content_difference
            && manifest_semantics_equal
            && non_manifest_content_exact
            && non_manifest_metadata_exact
            && non_manifest_record_exact
            && meta_record_exact
            && leaf_manifest_tracks_order
            && facade_manifest_tracks_order
            && !manifest_listing_tracks_order;
        let allowed = (member_order_exact
            && non_meta_content_exact
            && non_meta_metadata_exact
            && meta_metadata_exact
            && non_manifest_record_exact
            && meta_record_exact
            && content_differences.is_empty())
            || archive_member_order_and_manifest_listing;
        let decision = if archive_member_order_and_manifest_listing {
            "archive_member_order_and_manifest_listing"
        } else if content_differences.is_empty() {
            "byte_exact"
        } else {
            "unadmitted_content_or_metadata_delta"
        };
        let audit = json!({
            "status": "audited",
            "decision": decision,
            "allowed": allowed,
            "archive_comment_equal": true,
            "leaf_archive_comment_sha256": leaf_comment_sha256,
            "facade_archive_comment_sha256": facade_comment_sha256,
            "manifest_semantics_equal": manifest_semantics_equal,
            "leaf_manifest_semantics": manifest_semantics_json(&leaf_manifest_semantics),
            "facade_manifest_semantics": manifest_semantics_json(&facade_manifest_semantics),
            "member_order_exact": member_order_exact,
            "leaf_order": leaf_members.iter().map(|member| &member.name).collect::<Vec<_>>(),
            "facade_order": facade_members.iter().map(|member| &member.name).collect::<Vec<_>>(),
            "manifest_listing_tracks_order": manifest_listing_tracks_order,
            "leaf_manifest_tracks_order": leaf_manifest_tracks_order,
            "facade_manifest_tracks_order": facade_manifest_tracks_order,
            "archive_member_order_and_manifest_listing": archive_member_order_and_manifest_listing,
            "non_meta_content_exact": non_meta_content_exact,
            "non_manifest_content_exact": non_manifest_content_exact,
            "non_meta_stored_metadata_exact": non_meta_metadata_exact,
            "non_manifest_stored_metadata_exact": non_manifest_metadata_exact,
            "meta_stored_metadata_exact": meta_metadata_exact,
            "non_manifest_archive_record_exact": non_manifest_record_exact,
            "meta_archive_record_exact": meta_record_exact,
            "content_differences": content_differences,
            "leaf_member_count": leaf_members.len(),
            "facade_member_count": facade_members.len(),
            "members": member_rows,
        });
        assert!(allowed, "ODM archive audit failed: {audit}");
        audit
    }

    fn assert_source_pair(family: &str, native: &[u8], facade: &[u8]) {
        assert_eq!(
            native, facade,
            "{family} native and facade source bytes differ before editing"
        );
    }

    #[allow(
        clippy::too_many_arguments,
        reason = "the side receipt keeps all bounded triad and fidelity facts together"
    )]
    fn side(
        source: &[u8],
        source_unchanged: bool,
        output: &[u8],
        repeated: &[u8],
        forward: &[u8],
        inverse: &[u8],
        expected_changed: bool,
        changed: bool,
        patch_noop: bool,
        patch_base_match: bool,
        diagnostics_changed: bool,
        snapshot_bytes: Option<usize>,
        operations: Option<usize>,
        commits: Option<usize>,
    ) -> Value {
        let output_differs_from_source = output != source;
        let publication_repeat_exact = output == repeated;
        let forward_replay_exact = output == forward;
        let inverse_replay_restores_source = inverse == source;
        assert_eq!(
            changed, expected_changed,
            "commit changed flag disagrees with the route expectation"
        );
        assert_eq!(
            patch_noop, !expected_changed,
            "patch no-op flag disagrees with the route expectation"
        );
        assert!(patch_base_match, "patch base admission failed");
        assert_eq!(
            diagnostics_changed, expected_changed,
            "diagnostics changed flag disagrees with the route expectation"
        );
        assert!(source_unchanged, "source owner changed during commit");
        assert_eq!(
            output_differs_from_source, expected_changed,
            "published output/source difference disagrees with the route expectation"
        );
        assert!(
            publication_repeat_exact,
            "repeated publication bytes differ"
        );
        assert!(forward_replay_exact, "forward replay bytes differ");
        assert!(
            inverse_replay_restores_source,
            "inverse replay did not restore source bytes"
        );
        json!({
            "changed": changed,
            "patch_noop": patch_noop,
            "patch_base_match": patch_base_match,
            "diagnostics_changed": diagnostics_changed,
            "source_unchanged": source_unchanged,
            "output_differs_from_source": output_differs_from_source,
            "publication_repeat_exact": publication_repeat_exact,
            "forward_replay_exact": forward_replay_exact,
            "inverse_replay_restores_source": inverse_replay_restores_source,
            "source_bytes": source.len(),
            "output_bytes": output.len(),
            "output_sha256": digest(output),
            "diagnostic_snapshot_bytes": snapshot_bytes,
            "diagnostic_operations": operations,
            "diagnostic_commits": commits,
        })
    }

    fn phase_timings(
        iterations: usize,
        collect_allocations: bool,
        mut commit_only: impl FnMut(),
        mut commit_plus_diagnostics: impl FnMut(),
        mut publication: impl FnMut(),
    ) -> Value {
        fn allocation_json(sample: [usize; 5]) -> Value {
            json!({
                "allocations": sample[0],
                "deallocations": sample[1],
                "allocated_bytes": sample[2],
                "deallocated_bytes": sample[3],
                "reallocations": sample[4],
            })
        }

        fn measure(
            iterations: usize,
            collect_allocations: bool,
            mut operation: impl FnMut(),
        ) -> (u128, Vec<Value>) {
            let mut allocation_samples = if collect_allocations {
                Vec::with_capacity(iterations)
            } else {
                Vec::new()
            };
            let start = Instant::now();
            for _ in 0..iterations {
                allocation::begin();
                operation();
                if collect_allocations {
                    // Snapshot counters before constructing the JSON value;
                    // serialization bookkeeping must not be attributed to the
                    // operation that was just measured.
                    allocation_samples.push(allocation_json(allocation::end()));
                }
            }
            (
                start.elapsed().as_nanos() / iterations as u128,
                allocation_samples,
            )
        }

        let (commit_only_ns, commit_only_allocations) =
            measure(iterations, collect_allocations, &mut commit_only);
        let (diagnostics_ns, diagnostics_allocations) = measure(
            iterations,
            collect_allocations,
            &mut commit_plus_diagnostics,
        );
        let (publication_ns, publication_allocations) =
            measure(iterations, collect_allocations, &mut publication);
        json!({
            "iterations": iterations,
            "ns_per_operation": {
                "commit_only": commit_only_ns,
                "commit_plus_diagnostics": diagnostics_ns,
                "publication": publication_ns,
            },
            "operation_allocator_samples": {
                "commit_only": commit_only_allocations,
                "commit_plus_diagnostics": diagnostics_allocations,
                "publication": publication_allocations,
            },
        })
    }

    fn timing_contract() -> Value {
        json!({
            "status": "not_run",
            "setup_outside_clock": [
                "deterministic builder or tracked fixture read",
                "Package::from_bytes and File::open_bytes",
                "typed edit inputs, props and edit target construction",
            ],
            "clock_phases": [
                "commit_only",
                "commit_plus_diagnostics",
                "publication",
            ],
            "max_iterations": MAX_ITERATIONS,
            "allocation_boundary": {
                "facade_observation": "diagnostics.snapshot_bytes",
                "leaf_observation": "not exposed by the shared core Diagnostics",
                "process_allocator_capture":
                    "paired runner operation-scoped system allocator observer",
            },
        })
    }

    fn timing(
        mode: &str,
        iterations: usize,
        leaf_commit_only: impl FnMut(),
        leaf_commit_plus_diagnostics: impl FnMut(),
        leaf_publication: impl FnMut(),
        facade_commit_only: impl FnMut(),
        facade_commit_plus_diagnostics: impl FnMut(),
        facade_publication: impl FnMut(),
    ) -> Value {
        if mode == "timed" || mode == "allocator" {
            assert!(
                mode != "allocator" || allocation::ENABLED,
                "allocator mode requires the operation-scoped allocator build"
            );
            json!({
                "status": if mode == "allocator" {
                    "allocator_observed"
                } else {
                    "measured"
                },
                "elapsed_claimable": mode == "timed",
                "allocator_enabled": allocation::ENABLED,
                "leaf": phase_timings(
                    iterations,
                    mode == "allocator",
                    leaf_commit_only,
                    leaf_commit_plus_diagnostics,
                    leaf_publication,
                ),
                "facade": phase_timings(
                    iterations,
                    mode == "allocator",
                    facade_commit_only,
                    facade_commit_plus_diagnostics,
                    facade_publication,
                ),
            })
        } else {
            timing_contract()
        }
    }

    fn odc_changed(
        package: &litchi_odc::Package,
        definition: &litchi_odc::Definition,
    ) -> litchi_odc::Commit {
        let mut edit = package.edit();
        edit.replace_chart(definition).expect("ODC replace_chart");
        edit.commit().expect("ODC changed commit")
    }

    fn odc_noop(package: &litchi_odc::Package) -> litchi_odc::Commit {
        package.edit().commit().expect("ODC no-op commit")
    }

    fn odc_file_changed(
        file: &File,
        target: &TargetArg<'_>,
        properties: &[(&str, &str); 1],
    ) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set(target, properties)?;
            Ok(())
        })
        .expect("ODC facade changed commit")
    }

    fn odc_file_noop(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|_| Ok(())).expect("ODC facade no-op commit")
    }

    fn odc_probe(mode: &str, iterations: usize) -> Value {
        let package = litchi_odc::Package::new().expect("ODC deterministic builder");
        let file = open_bytes(package.as_bytes().to_vec()).expect("ODC facade builder");
        let source = package.as_bytes().to_vec();
        let mut changed_definition = package.chart().definition().expect("ODC definition");
        changed_definition.title = Some(litchi_odc::Text::new(CHANGED_TEXT));
        let target = TargetArg::parse("/").expect("ODC target");

        let changed = odc_changed(&package, &changed_definition);
        let noop = odc_noop(&package);
        let changed_output = changed.package().as_bytes().to_vec();
        let changed_repeat = changed.package().to_bytes().expect("ODC publish");
        let changed_forward = package
            .apply(changed.patch())
            .expect("ODC forward replay")
            .package()
            .as_bytes()
            .to_vec();
        let changed_inverse = changed
            .package()
            .apply(&changed.patch().inverse())
            .expect("ODC inverse replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_output = noop.package().as_bytes().to_vec();
        let noop_repeat = noop.package().to_bytes().expect("ODC no-op publish");
        let noop_forward = package
            .apply(noop.patch())
            .expect("ODC no-op replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_inverse = noop
            .package()
            .apply(&noop.patch().inverse())
            .expect("ODC no-op inverse")
            .package()
            .as_bytes()
            .to_vec();
        let leaf = json!({
            "changed": side(
                &source,
                package.as_bytes() == source,
                &changed_output,
                &changed_repeat,
                &changed_forward,
                &changed_inverse,
                true,
                changed.diagnostics().changed(),
                changed.patch().is_noop(),
                changed.patch().is_applicable_to(&package),
                changed.diagnostics().changed(),
                None::<usize>, None::<usize>, None::<usize>,
            ),
            "no_op": side(
                &source,
                package.as_bytes() == source,
                &noop_output,
                &noop_repeat,
                &noop_forward,
                &noop_inverse,
                false,
                noop.diagnostics().changed(),
                noop.patch().is_noop(),
                noop.patch().is_applicable_to(&package),
                noop.diagnostics().changed(),
                None::<usize>, None::<usize>, None::<usize>,
            ),
        });

        let file_changed = odc_file_changed(&file, &target, &ODC_PROPS);
        let file_noop = odc_file_noop(&file);
        let file_source = file.to_bytes().expect("ODC facade source");
        let file_changed_output = file_changed.file().to_bytes().expect("ODC facade publish");
        let file_changed_repeat = file_changed.file().to_bytes().expect("ODC facade repeat");
        let file_changed_forward = file_changed
            .patch()
            .apply(&file)
            .expect("ODC facade forward replay")
            .to_bytes()
            .expect("ODC facade forward bytes");
        let file_changed_inverse = file_changed
            .patch()
            .inverse()
            .apply(file_changed.file())
            .expect("ODC facade inverse replay")
            .to_bytes()
            .expect("ODC facade inverse bytes");
        let file_noop_output = file_noop.file().to_bytes().expect("ODC facade no-op");
        let file_noop_repeat = file_noop
            .file()
            .to_bytes()
            .expect("ODC facade no-op repeat");
        let file_noop_forward = file_noop
            .patch()
            .apply(&file)
            .expect("ODC facade no-op replay")
            .to_bytes()
            .expect("ODC facade no-op bytes");
        let file_noop_inverse = file_noop
            .patch()
            .inverse()
            .apply(file_noop.file())
            .expect("ODC facade no-op inverse")
            .to_bytes()
            .expect("ODC facade no-op inverse bytes");
        let facade = json!({
            "changed": side(
                &file_source,
                file.to_bytes().expect("ODC facade source repeat") == file_source,
                &file_changed_output,
                &file_changed_repeat,
                &file_changed_forward,
                &file_changed_inverse,
                true,
                file_changed.diagnostics().changed(),
                file_changed.patch().is_noop(),
                file_changed.patch().apply(&file).is_ok(),
                file_changed.diagnostics().changed(),
                Some(file_changed.diagnostics().snapshot_bytes()),
                Some(file_changed.diagnostics().operations()),
                Some(file_changed.diagnostics().commits()),
            ),
            "no_op": side(
                &file_source,
                file.to_bytes().expect("ODC facade no-op source") == file_source,
                &file_noop_output,
                &file_noop_repeat,
                &file_noop_forward,
                &file_noop_inverse,
                false,
                file_noop.diagnostics().changed(),
                file_noop.patch().is_noop(),
                file_noop.patch().apply(&file).is_ok(),
                file_noop.diagnostics().changed(),
                Some(file_noop.diagnostics().snapshot_bytes()),
                Some(file_noop.diagnostics().operations()),
                Some(file_noop.diagnostics().commits()),
            ),
        });

        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(odc_changed(&package, &changed_definition));
            },
            || {
                let commit = odc_changed(&package, &changed_definition);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odc_changed(&package, &changed_definition);
                black_box(commit.package().to_bytes().expect("ODC timed publish"));
            },
            || {
                black_box(odc_file_changed(&file, &target, &ODC_PROPS));
            },
            || {
                let commit = odc_file_changed(&file, &target, &ODC_PROPS);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odc_file_changed(&file, &target, &ODC_PROPS);
                black_box(commit.file().to_bytes().expect("ODC facade timed publish"));
            },
        );
        assert_source_pair("ODC", &source, &file_source);
        let leaf_title = litchi_odc::Package::from_bytes(&changed_output)
            .expect("ODC semantic leaf output")
            .chart()
            .definition()
            .expect("ODC semantic leaf definition")
            .title
            .map(|value| value.text);
        let facade_title = litchi_odc::Package::from_bytes(&file_changed_output)
            .expect("ODC semantic facade output")
            .chart()
            .definition()
            .expect("ODC semantic facade definition")
            .title
            .map(|value| value.text);
        family_result(
            FamilyIdentity {
                family: "odc",
                route: "odc_chart_title",
                source_kind: "deterministic_builder",
                source: "litchi_odc::Package::new",
            },
            &source,
            &file_source,
            &changed_output,
            &file_changed_output,
            leaf_title.as_deref() == Some(CHANGED_TEXT)
                && facade_title.as_deref() == Some(CHANGED_TEXT)
                && leaf_title == facade_title,
            true,
            Value::Null,
            leaf,
            facade,
            timing,
        )
    }

    fn odi_changed(
        package: &litchi_odi::Package,
        geometry: [Option<String>; 4],
    ) -> litchi_odi::Commit {
        let mut edit = package.edit();
        edit.set_geometry(
            0,
            geometry[0].clone(),
            geometry[1].clone(),
            geometry[2].clone(),
            geometry[3].clone(),
        )
        .expect("ODI set_geometry");
        edit.commit().expect("ODI changed commit")
    }

    fn odi_noop(package: &litchi_odi::Package) -> litchi_odi::Commit {
        package.edit().commit().expect("ODI no-op commit")
    }

    fn odi_file_changed(
        file: &File,
        target: &TargetArg<'_>,
        properties: &[(&str, &str); 4],
    ) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set(target, properties)?;
            Ok(())
        })
        .expect("ODI facade changed commit")
    }

    fn odi_file_noop(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|_| Ok(())).expect("ODI facade no-op commit")
    }

    fn odi_probe(mode: &str, iterations: usize) -> Value {
        let package = litchi_odi::Package::new().expect("ODI deterministic builder");
        let file = open_bytes(package.as_bytes().to_vec()).expect("ODI facade builder");
        let source = package.as_bytes().to_vec();
        let geometry = [
            Some("1cm".to_owned()),
            Some("2cm".to_owned()),
            Some("3cm".to_owned()),
            Some("4cm".to_owned()),
        ];
        let target = TargetArg::parse("/frame[1]").expect("ODI target");

        let changed = odi_changed(&package, geometry.clone());
        let noop = odi_noop(&package);
        let changed_output = changed.package().as_bytes().to_vec();
        let changed_repeat = changed.package().to_bytes().expect("ODI publish");
        let changed_forward = package
            .apply(changed.patch())
            .expect("ODI forward replay")
            .package()
            .as_bytes()
            .to_vec();
        let changed_inverse = changed
            .package()
            .apply(&changed.patch().inverse())
            .expect("ODI inverse replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_output = noop.package().as_bytes().to_vec();
        let noop_repeat = noop.package().to_bytes().expect("ODI no-op publish");
        let noop_forward = package
            .apply(noop.patch())
            .expect("ODI no-op replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_inverse = noop
            .package()
            .apply(&noop.patch().inverse())
            .expect("ODI no-op inverse")
            .package()
            .as_bytes()
            .to_vec();
        let leaf = json!({
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, true, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, false, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = odi_file_changed(&file, &target, &ODI_PROPS);
        let file_noop = odi_file_noop(&file);
        let file_source = file.to_bytes().expect("ODI facade source");
        let file_changed_output = file_changed.file().to_bytes().expect("ODI facade publish");
        let file_changed_repeat = file_changed.file().to_bytes().expect("ODI facade repeat");
        let file_changed_forward = file_changed
            .patch()
            .apply(&file)
            .expect("ODI facade forward replay")
            .to_bytes()
            .expect("ODI facade forward bytes");
        let file_changed_inverse = file_changed
            .patch()
            .inverse()
            .apply(file_changed.file())
            .expect("ODI facade inverse replay")
            .to_bytes()
            .expect("ODI facade inverse bytes");
        let file_noop_output = file_noop.file().to_bytes().expect("ODI facade no-op");
        let file_noop_repeat = file_noop
            .file()
            .to_bytes()
            .expect("ODI facade no-op repeat");
        let file_noop_forward = file_noop
            .patch()
            .apply(&file)
            .expect("ODI facade no-op replay")
            .to_bytes()
            .expect("ODI facade no-op bytes");
        let file_noop_inverse = file_noop
            .patch()
            .inverse()
            .apply(file_noop.file())
            .expect("ODI facade no-op inverse")
            .to_bytes()
            .expect("ODI facade no-op inverse bytes");
        let facade = json!({
            "changed": side(&file_source, file.to_bytes().expect("ODI facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, true, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("ODI facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, false, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let mut leaf_commit_geometry = vec![geometry.clone(); iterations];
        let mut leaf_diagnostics_geometry = vec![geometry.clone(); iterations];
        let mut leaf_publication_geometry = vec![geometry.clone(); iterations];
        let timing = timing(
            mode,
            iterations,
            || {
                let geometry = leaf_commit_geometry.pop().expect("ODI timed commit input");
                let _ = black_box(odi_changed(&package, geometry));
            },
            || {
                let geometry = leaf_diagnostics_geometry
                    .pop()
                    .expect("ODI timed diagnostics input");
                let commit = odi_changed(&package, geometry);
                black_box(commit.diagnostics().changed());
            },
            || {
                let geometry = leaf_publication_geometry
                    .pop()
                    .expect("ODI timed publication input");
                let commit = odi_changed(&package, geometry);
                black_box(commit.package().to_bytes().expect("ODI timed publish"));
            },
            || {
                black_box(odi_file_changed(&file, &target, &ODI_PROPS));
            },
            || {
                let commit = odi_file_changed(&file, &target, &ODI_PROPS);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odi_file_changed(&file, &target, &ODI_PROPS);
                black_box(commit.file().to_bytes().expect("ODI facade timed publish"));
            },
        );
        assert_source_pair("ODI", &source, &file_source);
        let geometry_of = |bytes: &[u8]| {
            let package = litchi_odi::Package::from_bytes(bytes).expect("ODI semantic output");
            let image = package.image();
            let frame = &image.frames()[0];
            [
                frame.x().map(str::to_owned),
                frame.y().map(str::to_owned),
                frame.width().map(str::to_owned),
                frame.height().map(str::to_owned),
            ]
        };
        let leaf_geometry = geometry_of(&changed_output);
        let facade_geometry = geometry_of(&file_changed_output);
        family_result(
            FamilyIdentity {
                family: "odi",
                route: "odi_frame_geometry",
                source_kind: "deterministic_builder",
                source: "litchi_odi::Package::new",
            },
            &source,
            &file_source,
            &changed_output,
            &file_changed_output,
            leaf_geometry == geometry
                && facade_geometry == geometry
                && leaf_geometry == facade_geometry,
            true,
            Value::Null,
            leaf,
            facade,
            timing,
        )
    }

    fn oth_changed(
        package: &litchi_oth::Package,
        position: Position,
        text: String,
    ) -> litchi_oth::Commit {
        let mut edit = package.edit();
        edit.set_paragraph_text(position, text)
            .expect("OTH set_paragraph_text");
        edit.commit().expect("OTH changed commit")
    }

    fn oth_noop(package: &litchi_oth::Package) -> litchi_oth::Commit {
        package.edit().commit().expect("OTH no-op commit")
    }

    fn oth_file_changed(
        file: &File,
        target: &TargetArg<'_>,
        properties: &[(&str, &str); 1],
    ) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set(target, properties)?;
            Ok(())
        })
        .expect("OTH facade changed commit")
    }

    fn oth_file_noop(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|_| Ok(())).expect("OTH facade no-op commit")
    }

    fn oth_probe(mode: &str, iterations: usize) -> Value {
        let package = litchi_oth::Builder::new()
            .paragraph(litchi_oth::paragraph::Paragraph::new("before"))
            .build()
            .expect("OTH deterministic builder");
        let file = open_bytes(package.as_bytes().to_vec()).expect("OTH facade builder");
        let source = package.as_bytes().to_vec();
        let changed_text = CHANGED_TEXT.to_owned();
        let position = Position::new(0);
        let target = TargetArg::parse("/body/p[1]").expect("OTH target");

        let changed = oth_changed(&package, position, changed_text.clone());
        let noop = oth_noop(&package);
        let changed_output = changed.package().as_bytes().to_vec();
        let changed_repeat = changed.package().to_bytes().expect("OTH publish");
        let changed_forward = package
            .apply(changed.patch())
            .expect("OTH forward replay")
            .package()
            .as_bytes()
            .to_vec();
        let changed_inverse = changed
            .package()
            .apply(&changed.patch().inverse())
            .expect("OTH inverse replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_output = noop.package().as_bytes().to_vec();
        let noop_repeat = noop.package().to_bytes().expect("OTH no-op publish");
        let noop_forward = package
            .apply(noop.patch())
            .expect("OTH no-op replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_inverse = noop
            .package()
            .apply(&noop.patch().inverse())
            .expect("OTH no-op inverse")
            .package()
            .as_bytes()
            .to_vec();
        let leaf = json!({
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, true, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, false, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = oth_file_changed(&file, &target, &OTH_PROPS);
        let file_noop = oth_file_noop(&file);
        let file_source = file.to_bytes().expect("OTH facade source");
        let file_changed_output = file_changed.file().to_bytes().expect("OTH facade publish");
        let file_changed_repeat = file_changed.file().to_bytes().expect("OTH facade repeat");
        let file_changed_forward = file_changed
            .patch()
            .apply(&file)
            .expect("OTH facade forward replay")
            .to_bytes()
            .expect("OTH facade forward bytes");
        let file_changed_inverse = file_changed
            .patch()
            .inverse()
            .apply(file_changed.file())
            .expect("OTH facade inverse replay")
            .to_bytes()
            .expect("OTH facade inverse bytes");
        let file_noop_output = file_noop.file().to_bytes().expect("OTH facade no-op");
        let file_noop_repeat = file_noop
            .file()
            .to_bytes()
            .expect("OTH facade no-op repeat");
        let file_noop_forward = file_noop
            .patch()
            .apply(&file)
            .expect("OTH facade no-op replay")
            .to_bytes()
            .expect("OTH facade no-op bytes");
        let file_noop_inverse = file_noop
            .patch()
            .inverse()
            .apply(file_noop.file())
            .expect("OTH facade no-op inverse")
            .to_bytes()
            .expect("OTH facade no-op inverse bytes");
        let facade = json!({
            "changed": side(&file_source, file.to_bytes().expect("OTH facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, true, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("OTH facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, false, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let mut leaf_commit_text = vec![changed_text.clone(); iterations];
        let mut leaf_diagnostics_text = vec![changed_text.clone(); iterations];
        let mut leaf_publication_text = vec![changed_text.clone(); iterations];
        let timing = timing(
            mode,
            iterations,
            || {
                let text = leaf_commit_text.pop().expect("OTH timed commit input");
                let _ = black_box(oth_changed(&package, position, text));
            },
            || {
                let text = leaf_diagnostics_text
                    .pop()
                    .expect("OTH timed diagnostics input");
                let commit = oth_changed(&package, position, text);
                black_box(commit.diagnostics().changed());
            },
            || {
                let text = leaf_publication_text
                    .pop()
                    .expect("OTH timed publication input");
                let commit = oth_changed(&package, position, text);
                black_box(commit.package().to_bytes().expect("OTH timed publish"));
            },
            || {
                black_box(oth_file_changed(&file, &target, &OTH_PROPS));
            },
            || {
                let commit = oth_file_changed(&file, &target, &OTH_PROPS);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = oth_file_changed(&file, &target, &OTH_PROPS);
                black_box(commit.file().to_bytes().expect("OTH facade timed publish"));
            },
        );
        let leaf_text = litchi_oth::Package::from_bytes(&changed_output)
            .expect("OTH semantic leaf output")
            .document()
            .text_body()
            .expect("OTH semantic leaf body")
            .paragraphs()[0]
            .text()
            .to_owned();
        let facade_text = litchi_oth::Package::from_bytes(&file_changed_output)
            .expect("OTH semantic facade output")
            .document()
            .text_body()
            .expect("OTH semantic facade body")
            .paragraphs()[0]
            .text()
            .to_owned();
        assert_source_pair("OTH", &source, &file_source);
        family_result(
            FamilyIdentity {
                family: "oth",
                route: "oth_paragraph_text",
                source_kind: "deterministic_builder",
                source: "litchi_oth::Builder::paragraph(before)",
            },
            &source,
            &file_source,
            &changed_output,
            &file_changed_output,
            leaf_text == CHANGED_TEXT && facade_text == CHANGED_TEXT && leaf_text == facade_text,
            true,
            Value::Null,
            leaf,
            facade,
            timing,
        )
    }

    fn odm_source() -> litchi_odm::Package {
        // The ODM builder is deterministic but intentionally omits meta.xml.
        // Add the canonical timestamp-free metadata skeleton through the
        // shared bounded ODF writer so the title transaction has a real
        // metadata part without importing an external corpus archive.
        let built = litchi_odm::Builder::new()
            .build()
            .expect("ODM deterministic builder");
        let mut writer = PackageWriter::new(StreamingArchiveLimits::default());
        writer
            .set_mimetype("application/vnd.oasis.opendocument.text-master")
            .expect("ODM mimetype");
        writer
            .add_file("meta.xml", Structure::default_meta_xml().as_bytes())
            .expect("ODM metadata");
        writer
            .add_file("content.xml", built.document().content_xml().as_bytes())
            .expect("ODM content");
        litchi_odm::Package::from_bytes(writer.finish().expect("ODM package").as_slice())
            .expect("ODM deterministic package")
    }

    fn odm_changed(
        package: &litchi_odm::Package,
        title: String,
    ) -> litchi_odm::transaction::Commit {
        let mut edit = package.edit();
        edit.set_title(title).expect("ODM set_title");
        edit.commit().expect("ODM changed commit")
    }

    fn odm_noop(package: &litchi_odm::Package) -> litchi_odm::transaction::Commit {
        package.edit().commit().expect("ODM no-op commit")
    }

    fn odm_file_changed(
        file: &File,
        target: &TargetArg<'_>,
        properties: &[(&str, &str); 1],
    ) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set(target, properties)?;
            Ok(())
        })
        .expect("ODM facade changed commit")
    }

    fn odm_file_noop(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|_| Ok(())).expect("ODM facade no-op commit")
    }

    fn odm_probe(mode: &str, iterations: usize) -> Value {
        let package = odm_source();
        let source = package.as_bytes().to_vec();
        let file = open_bytes(source.clone()).expect("ODM facade builder");
        let changed_title = CHANGED_TEXT.to_owned();
        let target = TargetArg::parse("/metadata").expect("ODM target");

        let changed = odm_changed(&package, changed_title.clone());
        let noop = odm_noop(&package);
        let changed_output = changed.package().as_bytes().to_vec();
        let changed_repeat = changed.package().to_bytes().expect("ODM publish");
        let changed_forward = package
            .apply(changed.patch())
            .expect("ODM forward replay")
            .package()
            .as_bytes()
            .to_vec();
        let changed_inverse = changed
            .package()
            .apply(&changed.patch().inverse())
            .expect("ODM inverse replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_output = noop.package().as_bytes().to_vec();
        let noop_repeat = noop.package().to_bytes().expect("ODM no-op publish");
        let noop_forward = package
            .apply(noop.patch())
            .expect("ODM no-op replay")
            .package()
            .as_bytes()
            .to_vec();
        let noop_inverse = noop
            .package()
            .apply(&noop.patch().inverse())
            .expect("ODM no-op inverse")
            .package()
            .as_bytes()
            .to_vec();
        let leaf = json!({
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, true, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, false, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = odm_file_changed(&file, &target, &ODM_PROPS);
        let file_noop = odm_file_noop(&file);
        let file_source = file.to_bytes().expect("ODM facade source");
        let file_changed_output = file_changed.file().to_bytes().expect("ODM facade publish");
        let file_changed_repeat = file_changed.file().to_bytes().expect("ODM facade repeat");
        let file_changed_forward = file_changed
            .patch()
            .apply(&file)
            .expect("ODM facade forward replay")
            .to_bytes()
            .expect("ODM facade forward bytes");
        let file_changed_inverse = file_changed
            .patch()
            .inverse()
            .apply(file_changed.file())
            .expect("ODM facade inverse replay")
            .to_bytes()
            .expect("ODM facade inverse bytes");
        let file_noop_output = file_noop.file().to_bytes().expect("ODM facade no-op");
        let file_noop_repeat = file_noop
            .file()
            .to_bytes()
            .expect("ODM facade no-op repeat");
        let file_noop_forward = file_noop
            .patch()
            .apply(&file)
            .expect("ODM facade no-op replay")
            .to_bytes()
            .expect("ODM facade no-op bytes");
        let file_noop_inverse = file_noop
            .patch()
            .inverse()
            .apply(file_noop.file())
            .expect("ODM facade no-op inverse")
            .to_bytes()
            .expect("ODM facade no-op inverse bytes");
        let facade = json!({
            "changed": side(&file_source, file.to_bytes().expect("ODM facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, true, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("ODM facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, false, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let mut leaf_commit_title = vec![changed_title.clone(); iterations];
        let mut leaf_diagnostics_title = vec![changed_title.clone(); iterations];
        let mut leaf_publication_title = vec![changed_title.clone(); iterations];
        let timing = timing(
            mode,
            iterations,
            || {
                let title = leaf_commit_title.pop().expect("ODM timed commit input");
                let _ = black_box(odm_changed(&package, title));
            },
            || {
                let title = leaf_diagnostics_title
                    .pop()
                    .expect("ODM timed diagnostics input");
                let commit = odm_changed(&package, title);
                black_box(commit.diagnostics().changed());
            },
            || {
                let title = leaf_publication_title
                    .pop()
                    .expect("ODM timed publication input");
                let commit = odm_changed(&package, title);
                black_box(commit.package().to_bytes().expect("ODM timed publish"));
            },
            || {
                black_box(odm_file_changed(&file, &target, &ODM_PROPS));
            },
            || {
                let commit = odm_file_changed(&file, &target, &ODM_PROPS);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odm_file_changed(&file, &target, &ODM_PROPS);
                black_box(commit.file().to_bytes().expect("ODM facade timed publish"));
            },
        );
        let leaf_title = litchi_odm::Package::from_bytes(&changed_output)
            .expect("ODM semantic leaf output")
            .document()
            .title()
            .map(str::to_owned);
        let facade_title = litchi_odm::Package::from_bytes(&file_changed_output)
            .expect("ODM semantic facade output")
            .document()
            .title()
            .map(str::to_owned);
        assert_source_pair("ODM", &source, &file_source);
        let archive_audit = archive_pair(&changed_output, &file_changed_output);
        family_result(
            FamilyIdentity {
                family: "odm",
                route: "odm_title",
                source_kind: "deterministic_builder",
                source: "litchi_odm::Builder::new + timestamp-free Structure::default_meta_xml",
            },
            &source,
            &file_source,
            &changed_output,
            &file_changed_output,
            leaf_title.as_deref() == Some(CHANGED_TEXT)
                && facade_title.as_deref() == Some(CHANGED_TEXT)
                && leaf_title == facade_title,
            false,
            archive_audit,
            leaf,
            facade,
            timing,
        )
    }

    struct FamilyIdentity<'a> {
        family: &'a str,
        route: &'a str,
        source_kind: &'a str,
        source: &'a str,
    }

    fn family_result(
        identity: FamilyIdentity<'_>,
        source_bytes: &[u8],
        facade_source_bytes: &[u8],
        leaf_output: &[u8],
        facade_output: &[u8],
        semantic_equal: bool,
        require_changed_byte_identity: bool,
        archive_audit: Value,
        leaf: Value,
        facade: Value,
        timing: Value,
    ) -> Value {
        let FamilyIdentity {
            family,
            route,
            source_kind,
            source,
        } = identity;
        assert_eq!(
            source_bytes, facade_source_bytes,
            "{family} native/facade source bytes differ"
        );
        let changed_output_exact = leaf_output == facade_output;
        assert!(
            semantic_equal,
            "{family} changed outputs failed semantic proof"
        );
        if require_changed_byte_identity {
            assert!(
                changed_output_exact,
                "{family} changed native/facade output bytes differ"
            );
        }
        let (leaf_route, facade_route, changed_marker, no_op) = match family {
            "odc" => (
                "Package::edit -> replace_chart(title) -> commit",
                "File::edit -> set(/, title) -> commit",
                CHANGED_TEXT,
                "Package::edit/File::edit with an empty closure",
            ),
            "odi" => (
                "Package::edit -> set_geometry(frame 0) -> commit",
                "File::edit -> set(/frame[1], geometry) -> commit",
                "1cm,2cm,3cm,4cm frame geometry",
                "Package::edit/File::edit with an empty closure",
            ),
            "oth" => (
                "Package::edit -> set_paragraph_text(position 0) -> commit",
                "File::edit -> set(/body/p[1], text) -> commit",
                CHANGED_TEXT,
                "Package::edit/File::edit with an empty closure",
            ),
            "odm" => (
                "Package::edit -> set_title -> commit",
                "File::edit -> set(/metadata, title) -> commit",
                CHANGED_TEXT,
                "Package::edit/File::edit with an empty closure",
            ),
            _ => panic!("unknown shared-triad family {family}"),
        };
        json!({
            "family": family,
            "route": route,
            "descriptor": {
                "id": route,
                "family": family,
                "source": source,
                "leaf": leaf_route,
                "facade": facade_route,
                "changed_marker": changed_marker,
                "no_op": no_op,
            },
            "source_kind": source_kind,
            "source": source,
            "source_bytes": source_bytes.len(),
            "source_sha256": digest(source_bytes),
            "facade_source_bytes": facade_source_bytes.len(),
            "facade_source_sha256": digest(facade_source_bytes),
            "source_pair_exact": true,
            "leaf": leaf,
            "facade": facade,
            "paired": {
                "changed_output_exact": changed_output_exact,
                "changed_output_sha256": digest(leaf_output),
                "facade_changed_output_sha256": digest(facade_output),
                "semantic_equality": semantic_equal,
                "archive_audit": archive_audit,
            },
            "timing": timing,
        })
    }

    fn args() -> (String, Option<String>, usize) {
        let mut mode = "validate".to_owned();
        let mut family = None;
        let mut iterations = 1;
        let mut arguments = std::env::args().skip(1);
        while let Some(argument) = arguments.next() {
            match argument.as_str() {
                "--mode" => mode = arguments.next().expect("--mode needs a value"),
                "--family" => family = Some(arguments.next().expect("--family needs a value")),
                "--iterations" => {
                    iterations = arguments
                        .next()
                        .expect("--iterations needs a value")
                        .parse()
                        .expect("--iterations is an integer")
                },
                unknown => {
                    panic!("unknown argument {unknown}; use --mode, --family or --iterations")
                },
            }
        }
        assert!(
            mode == "validate" || mode == "timed" || mode == "allocator",
            "mode must be validate, timed or allocator"
        );
        assert!(
            (1..=MAX_ITERATIONS).contains(&iterations),
            "iterations must be 1..={MAX_ITERATIONS}"
        );
        if let Some(name) = &family {
            assert!(
                matches!(name.as_str(), "odc" | "odi" | "oth" | "odm" | "all"),
                "unknown family {name}"
            );
        }
        (mode, family, iterations)
    }

    pub(super) fn run() {
        let (mode, requested_family, iterations) = args();
        let selected = |name: &str| {
            requested_family
                .as_deref()
                .is_none_or(|family| family == "all" || family == name)
        };
        let mut rows = Vec::new();
        if selected("odc") {
            rows.push(odc_probe(&mode, iterations));
        }
        if selected("odi") {
            rows.push(odi_probe(&mode, iterations));
        }
        if selected("oth") {
            rows.push(oth_probe(&mode, iterations));
        }
        if selected("odm") {
            rows.push(odm_probe(&mode, iterations));
        }
        println!(
            "{}",
            serde_json::to_string_pretty(&json!({
                "schema": "litchi.odf.shared-triad-qualification.v1",
                "mode": mode,
                "iterations": iterations,
                "families": rows,
                "claim_scope": "correctness only unless mode=timed; no acceptance claim",
            }))
            .expect("qualification JSON")
        );
    }
}

fn main() {
    #[cfg(all(feature = "odc", feature = "odi", feature = "oth", feature = "odm"))]
    probe::run();
    #[cfg(not(all(feature = "odc", feature = "odi", feature = "oth", feature = "odm")))]
    panic!("enable odc, odi, oth and odm to run the shared-triad qualification probe");
}
