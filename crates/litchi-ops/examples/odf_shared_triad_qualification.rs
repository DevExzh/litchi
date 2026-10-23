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
    use std::time::Instant;

    use litchi_core::Position;
    use litchi_odf_common::core::{PackageWriter, StreamingArchiveLimits, Structure};
    use litchi_ops::{File, open_bytes, props};
    use serde_json::{Value, json};
    use sha2::{Digest as _, Sha256};

    const MAX_ITERATIONS: usize = 128;
    const CHANGED_TEXT: &str = "shared-triad qualification";

    fn digest(bytes: &[u8]) -> String {
        let mut output = String::with_capacity(64);
        for byte in Sha256::digest(bytes) {
            write!(output, "{byte:02x}").expect("digest formatting");
        }
        output
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
        changed: bool,
        patch_noop: bool,
        patch_base_match: bool,
        diagnostics_changed: bool,
        snapshot_bytes: Option<usize>,
        operations: Option<usize>,
        commits: Option<usize>,
    ) -> Value {
        json!({
            "changed": changed,
            "patch_noop": patch_noop,
            "patch_base_match": patch_base_match,
            "diagnostics_changed": diagnostics_changed,
            "source_unchanged": source_unchanged,
            "output_differs_from_source": output != source,
            "publication_repeat_exact": output == repeated,
            "forward_replay_exact": output == forward,
            "inverse_replay_restores_source": inverse == source,
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
                "tracked fixture read or deterministic builder",
                "native Package::from_bytes/open and File::open_bytes",
                "props and edit target construction",
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

    fn odc_changed(package: &litchi_odc::Package) -> litchi_odc::Commit {
        let mut definition = package.chart().definition().expect("ODC definition");
        definition.title = Some(litchi_odc::Text::new(CHANGED_TEXT));
        let mut edit = package.edit();
        edit.replace_chart(&definition).expect("ODC replace_chart");
        edit.commit().expect("ODC changed commit")
    }

    fn odc_noop(package: &litchi_odc::Package) -> litchi_odc::Commit {
        package.edit().commit().expect("ODC no-op commit")
    }

    fn odc_file_changed(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set("/", props! { "title" => CHANGED_TEXT })?;
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

        let changed = odc_changed(&package);
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
                noop.diagnostics().changed(),
                noop.patch().is_noop(),
                noop.patch().is_applicable_to(&package),
                noop.diagnostics().changed(),
                None::<usize>, None::<usize>, None::<usize>,
            ),
        });

        let file_changed = odc_file_changed(&file);
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
                let _ = black_box(odc_changed(&package));
            },
            || {
                let commit = odc_changed(&package);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odc_changed(&package);
                black_box(commit.package().to_bytes().expect("ODC timed publish"));
            },
            || {
                black_box(odc_file_changed(&file));
            },
            || {
                let commit = odc_file_changed(&file);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odc_file_changed(&file);
                black_box(commit.file().to_bytes().expect("ODC facade timed publish"));
            },
        );
        family_result(
            "odc",
            "deterministic_builder",
            "litchi_odc::Package::new",
            &source,
            &changed_output,
            &file_changed_output,
            true,
            leaf,
            facade,
            timing,
        )
    }

    fn odi_changed(package: &litchi_odi::Package) -> litchi_odi::Commit {
        let mut edit = package.edit();
        edit.set_geometry(
            0,
            Some("1cm".into()),
            Some("2cm".into()),
            Some("3cm".into()),
            Some("4cm".into()),
        )
        .expect("ODI set_geometry");
        edit.commit().expect("ODI changed commit")
    }

    fn odi_noop(package: &litchi_odi::Package) -> litchi_odi::Commit {
        package.edit().commit().expect("ODI no-op commit")
    }

    fn odi_file_changed(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set(
                "/frame[1]",
                props! { "x" => "1cm", "y" => "2cm", "width" => "3cm", "height" => "4cm" },
            )?;
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

        let changed = odi_changed(&package);
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
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = odi_file_changed(&file);
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
            "changed": side(&file_source, file.to_bytes().expect("ODI facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("ODI facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(odi_changed(&package));
            },
            || {
                let commit = odi_changed(&package);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odi_changed(&package);
                black_box(commit.package().to_bytes().expect("ODI timed publish"));
            },
            || {
                black_box(odi_file_changed(&file));
            },
            || {
                let commit = odi_file_changed(&file);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odi_file_changed(&file);
                black_box(commit.file().to_bytes().expect("ODI facade timed publish"));
            },
        );
        family_result(
            "odi",
            "deterministic_builder",
            "litchi_odi::Package::new",
            &source,
            &changed_output,
            &file_changed_output,
            true,
            leaf,
            facade,
            timing,
        )
    }

    fn oth_changed(package: &litchi_oth::Package) -> litchi_oth::Commit {
        let mut edit = package.edit();
        edit.set_paragraph_text(Position::new(0), CHANGED_TEXT)
            .expect("OTH set_paragraph_text");
        edit.commit().expect("OTH changed commit")
    }

    fn oth_noop(package: &litchi_oth::Package) -> litchi_oth::Commit {
        package.edit().commit().expect("OTH no-op commit")
    }

    fn oth_file_changed(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set("/body/p[1]", props! { "text" => CHANGED_TEXT })?;
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

        let changed = oth_changed(&package);
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
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = oth_file_changed(&file);
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
            "changed": side(&file_source, file.to_bytes().expect("OTH facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("OTH facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(oth_changed(&package));
            },
            || {
                let commit = oth_changed(&package);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = oth_changed(&package);
                black_box(commit.package().to_bytes().expect("OTH timed publish"));
            },
            || {
                black_box(oth_file_changed(&file));
            },
            || {
                let commit = oth_file_changed(&file);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = oth_file_changed(&file);
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
        family_result(
            "oth",
            "deterministic_builder",
            "litchi_oth::Builder::paragraph",
            &source,
            &changed_output,
            &file_changed_output,
            leaf_text == CHANGED_TEXT && facade_text == CHANGED_TEXT && leaf_text == facade_text,
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
            .add_file("content.xml", built.document().content_xml().as_bytes())
            .expect("ODM content");
        writer
            .add_file("meta.xml", Structure::default_meta_xml().as_bytes())
            .expect("ODM metadata");
        litchi_odm::Package::from_bytes(writer.finish().expect("ODM package").as_slice())
            .expect("ODM deterministic package")
    }

    fn odm_changed(package: &litchi_odm::Package) -> litchi_odm::transaction::Commit {
        let mut edit = package.edit();
        edit.set_title(CHANGED_TEXT).expect("ODM set_title");
        edit.commit().expect("ODM changed commit")
    }

    fn odm_noop(package: &litchi_odm::Package) -> litchi_odm::transaction::Commit {
        package.edit().commit().expect("ODM no-op commit")
    }

    fn odm_file_changed(file: &File) -> litchi_ops::transaction::Commit {
        file.edit(|edit| {
            edit.set("/metadata", props! { "title" => CHANGED_TEXT })?;
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

        let changed = odm_changed(&package);
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
            "changed": side(&source, package.as_bytes() == source, &changed_output, &changed_repeat, &changed_forward, &changed_inverse, changed.diagnostics().changed(), changed.patch().is_noop(), changed.patch().is_applicable_to(&package), changed.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
            "no_op": side(&source, package.as_bytes() == source, &noop_output, &noop_repeat, &noop_forward, &noop_inverse, noop.diagnostics().changed(), noop.patch().is_noop(), noop.patch().is_applicable_to(&package), noop.diagnostics().changed(), None::<usize>, None::<usize>, None::<usize>),
        });

        let file_changed = odm_file_changed(&file);
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
            "changed": side(&file_source, file.to_bytes().expect("ODM facade source repeat") == file_source, &file_changed_output, &file_changed_repeat, &file_changed_forward, &file_changed_inverse, file_changed.diagnostics().changed(), file_changed.patch().is_noop(), file_changed.patch().apply(&file).is_ok(), file_changed.diagnostics().changed(), Some(file_changed.diagnostics().snapshot_bytes()), Some(file_changed.diagnostics().operations()), Some(file_changed.diagnostics().commits())),
            "no_op": side(&file_source, file.to_bytes().expect("ODM facade no-op source") == file_source, &file_noop_output, &file_noop_repeat, &file_noop_forward, &file_noop_inverse, file_noop.diagnostics().changed(), file_noop.patch().is_noop(), file_noop.patch().apply(&file).is_ok(), file_noop.diagnostics().changed(), Some(file_noop.diagnostics().snapshot_bytes()), Some(file_noop.diagnostics().operations()), Some(file_noop.diagnostics().commits())),
        });
        let timing = timing(
            mode,
            iterations,
            || {
                let _ = black_box(odm_changed(&package));
            },
            || {
                let commit = odm_changed(&package);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odm_changed(&package);
                black_box(commit.package().to_bytes().expect("ODM timed publish"));
            },
            || {
                black_box(odm_file_changed(&file));
            },
            || {
                let commit = odm_file_changed(&file);
                black_box(commit.diagnostics().changed());
            },
            || {
                let commit = odm_file_changed(&file);
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
        family_result(
            "odm",
            "deterministic_builder",
            "litchi_odm::Builder::new",
            &source,
            &changed_output,
            &file_changed_output,
            leaf_title.as_deref() == Some(CHANGED_TEXT)
                && facade_title.as_deref() == Some(CHANGED_TEXT)
                && leaf_title == facade_title,
            leaf,
            facade,
            timing,
        )
    }

    fn family_result(
        family: &str,
        source_kind: &str,
        source: &str,
        source_bytes: &[u8],
        leaf_output: &[u8],
        facade_output: &[u8],
        semantic_equal: bool,
        leaf: Value,
        facade: Value,
        timing: Value,
    ) -> Value {
        json!({
            "family": family,
            "source_kind": source_kind,
            "source": source,
            "source_bytes": source_bytes.len(),
            "source_sha256": digest(source_bytes),
            "leaf": leaf,
            "facade": facade,
            "paired": {
                "changed_output_exact": leaf_output == facade_output,
                "changed_output_sha256": digest(leaf_output),
                "facade_changed_output_sha256": digest(facade_output),
                "semantic_equality": semantic_equal,
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
