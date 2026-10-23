//! Correctness and optional bounded phase timing for the ODF shared-triad lane.
//!
//! This probe is intentionally separate from the W8 109-route packet.  It
//! pairs the native owner and the `litchi_ops::File` route for the four ODF
//! object/text families changed by 0362usd.  The default `validate` mode is
//! untimed and proves source identity, exact no-ops, patch replay/inverse,
//! diagnostics and repeated publication.  `timed` is an execution aid for a
//! later quiet-host qualification: fixture I/O and owner opening stay outside
//! the measured closures, and the iteration count is bounded.

#[cfg(all(feature = "odc", feature = "odi", feature = "oth", feature = "odm"))]
mod probe {
    use std::fmt::Write as _;
    use std::hint::black_box;
    use std::io::{Cursor, Read};
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_odf_common::core::{PackageWriter, StreamingArchiveLimits, Structure};
    use litchi_ops::{File, TargetArg, open_bytes};
    use serde_json::{Value, json};
    use sha2::{Digest as _, Sha256};
    use zip::ZipArchive;

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
        last_modified: Option<String>,
        extra_sha256: Option<String>,
        comment: String,
        unix_mode: Option<u32>,
        version_made_by: (u8, u8),
        encrypted: bool,
        content_text: Option<String>,
    }

    fn archive_members(bytes: &[u8]) -> (String, Vec<ArchiveMember>) {
        let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("ODF ZIP archive");
        let archive_comment = digest(archive.comment());
        let mut members = Vec::with_capacity(archive.len());
        for index in 0..archive.len() {
            let mut member = archive.by_index(index).expect("ODF ZIP member");
            let name = member.name().to_owned();
            let name_raw_sha256 = digest(member.name_raw());
            let compression = format!("{:?}", member.compression());
            let compressed_bytes = member.compressed_size();
            let crc32 = member.crc32();
            let last_modified = member.last_modified().map(|value| format!("{value:?}"));
            let extra_sha256 = member.extra_data().map(digest);
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
            "last_modified": member.last_modified,
            "extra_sha256": member.extra_sha256,
            "comment": member.comment,
            "unix_mode": member.unix_mode,
            "version_made_by": member.version_made_by,
            "encrypted": member.encrypted,
            "content_text": member.content_text,
        })
    }

    fn manifest_paths(member: &ArchiveMember) -> Vec<String> {
        member
            .content_text
            .as_deref()
            .unwrap_or_default()
            .split("<manifest:file-entry ")
            .skip(1)
            .filter_map(|entry| {
                entry
                    .split_once("manifest:full-path=\"")
                    .and_then(|(_, value)| value.split_once('"'))
                    .map(|(path, _)| path.to_owned())
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

    fn manifest_tracks_order(member: &ArchiveMember, members: &[ArchiveMember]) -> bool {
        let manifest_payloads = manifest_paths(member)
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
        let only_meta_content_difference =
            content_differences.iter().all(|name| name == "meta.xml");
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
        let manifest_listing_tracks_order =
            manifest_paths(leaf_manifest) == manifest_paths(facade_manifest);
        let leaf_manifest_tracks_order = manifest_tracks_order(leaf_manifest, &leaf_members);
        let facade_manifest_tracks_order = manifest_tracks_order(facade_manifest, &facade_members);
        let archive_member_order_and_manifest_listing = !member_order_exact
            && manifest_content_difference
            && non_manifest_content_exact
            && non_manifest_metadata_exact
            && non_manifest_record_exact
            && meta_record_exact
            && leaf_manifest_tracks_order
            && facade_manifest_tracks_order
            && !manifest_listing_tracks_order;
        let allowed = (member_order_exact
            && leaf_comment == facade_comment
            && non_meta_content_exact
            && non_meta_metadata_exact
            && meta_metadata_exact
            && non_manifest_record_exact
            && meta_record_exact
            && only_meta_content_difference)
            || archive_member_order_and_manifest_listing;
        let decision = if archive_member_order_and_manifest_listing {
            "archive_member_order_and_manifest_listing"
        } else if content_differences.is_empty() {
            "byte_exact"
        } else {
            "metadata_member_only"
        };
        let audit = json!({
            "status": "audited",
            "decision": decision,
            "allowed": allowed,
            "archive_comment_equal": leaf_comment == facade_comment,
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
        mut commit_only: impl FnMut(),
        mut commit_plus_diagnostics: impl FnMut(),
        mut publication: impl FnMut(),
    ) -> Value {
        fn measure(iterations: usize, mut operation: impl FnMut()) -> u128 {
            let start = Instant::now();
            for _ in 0..iterations {
                operation();
            }
            start.elapsed().as_nanos() / iterations as u128
        }

        json!({
            "iterations": iterations,
            "ns_per_operation": {
                "commit_only": measure(iterations, &mut commit_only),
                "commit_plus_diagnostics": measure(iterations, &mut commit_plus_diagnostics),
                "publication": measure(iterations, &mut publication),
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
                "process_allocator_capture": "external runner only",
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
        if mode == "timed" {
            json!({
                "status": "measured",
                "leaf": phase_timings(
                    iterations,
                    leaf_commit_only,
                    leaf_commit_plus_diagnostics,
                    leaf_publication,
                ),
                "facade": phase_timings(
                    iterations,
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
            "odc",
            "odc_chart_title",
            "deterministic_builder",
            "litchi_odc::Package::new",
            &source,
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
        geometry: &[Option<String>; 4],
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

        let changed = odi_changed(&package, &geometry);
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
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(odi_changed(&package, &geometry));
            },
            || {
                let commit = odi_changed(&package, &geometry);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odi_changed(&package, &geometry);
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
            "odi",
            "odi_frame_geometry",
            "deterministic_builder",
            "litchi_odi::Package::new",
            &source,
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
        text: &str,
    ) -> litchi_oth::Commit {
        let mut edit = package.edit();
        edit.set_paragraph_text(position, text.to_owned())
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

        let changed = oth_changed(&package, position, &changed_text);
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
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(oth_changed(&package, position, &changed_text));
            },
            || {
                let commit = oth_changed(&package, position, &changed_text);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = oth_changed(&package, position, &changed_text);
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
            "oth",
            "oth_paragraph_text",
            "deterministic_builder",
            "litchi_oth::Builder::paragraph(before)",
            &source,
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

    fn odm_changed(package: &litchi_odm::Package, title: &str) -> litchi_odm::transaction::Commit {
        let mut edit = package.edit();
        edit.set_title(title.to_owned()).expect("ODM set_title");
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

        let changed = odm_changed(&package, &changed_title);
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
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(odm_changed(&package, &changed_title));
            },
            || {
                let commit = odm_changed(&package, &changed_title);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odm_changed(&package, &changed_title);
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
            "odm",
            "odm_title",
            "deterministic_builder",
            "litchi_odm::Builder::new + timestamp-free Structure::default_meta_xml",
            &source,
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

    fn family_result(
        family: &str,
        route: &str,
        source_kind: &str,
        source: &str,
        source_bytes: &[u8],
        leaf_output: &[u8],
        facade_output: &[u8],
        semantic_equal: bool,
        require_changed_byte_identity: bool,
        archive_audit: Value,
        leaf: Value,
        facade: Value,
        timing: Value,
    ) -> Value {
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
        json!({
            "family": family,
            "route": route,
            "source_kind": source_kind,
            "source": source,
            "source_bytes": source_bytes.len(),
            "source_sha256": digest(source_bytes),
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
            mode == "validate" || mode == "timed",
            "mode must be validate or timed"
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
