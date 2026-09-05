//! Concurrency and snapshot-isolation coverage for Pages body-table hidden axes.
//!
//! The source fixture lives in `body_table_hidden_axes.rs`; embedding that
//! module keeps this focused target on the same strict native graph while
//! avoiding a second fixture implementation.  The tests below deliberately
//! use only the public `Package` transaction surface.

mod fixture {
    include!("body_table_hidden_axes.rs");

    use std::sync::Barrier;

    const WORKERS: usize = 8;
    const READ_ROUNDS: usize = 12;

    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>(_: &T) {}

    #[test]
    fn snapshots_and_transaction_values_are_send_sync_and_debug() -> TestResult {
        let package = Package::from_bytes(&normal_package()?)?;
        let expected = HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?;
        assert_send_sync_debug(&package);
        assert_send_sync_debug(&expected);
        assert_send_sync_debug(&AxisIndex::row(1));
        assert_send_sync_debug(&BodyTableSelector::name("Revenue"));
        assert_send_sync_debug(&BodyTableHiddenAxesLimitKind::WireBytes);
        assert_send_sync_debug(&Error::PatchConflict);

        let edit = package.edit_body_table_hidden_axes(0usize)?;
        assert_send_sync_debug(&edit);
        let commit = edit.set(expected).commit()?;
        assert_send_sync_debug(&commit);
        assert_send_sync_debug(commit.patch());
        assert_send_sync_debug(commit.diagnostics());
        Ok(())
    }

    #[test]
    fn concurrent_first_reads_are_deterministic_without_shared_mutation() -> TestResult {
        let source = normal_package()?;
        let package = Arc::new(Package::from_bytes(&source)?);
        let expected = HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?;
        let start = Arc::new(Barrier::new(WORKERS));
        let mut handles = Vec::with_capacity(WORKERS);

        for worker in 0..WORKERS {
            let package = Arc::clone(&package);
            let start = Arc::clone(&start);
            let expected = expected.clone();
            handles.push(thread::spawn(move || {
                start.wait();
                for round in 0..READ_ROUNDS {
                    let by_position = package
                        .body_table_hidden_axes(0usize)
                        .expect("concurrent position read succeeds");
                    assert_eq!(by_position, expected, "worker {worker}, round {round}");
                    let by_name = package
                        .body_table_hidden_axes(BodyTableSelector::name("Revenue"))
                        .expect("concurrent name read succeeds");
                    assert_eq!(by_name, expected, "worker {worker}, round {round}");
                    assert_eq!(
                        package
                            .body_table_hidden_axes(1usize)
                            .expect("concurrent absent-state read succeeds"),
                        HiddenAxes::empty()
                    );
                }
            }));
        }

        for handle in handles {
            handle.join().expect("read worker must not panic");
        }
        assert_eq!(package.exact_bytes(), source);
        assert_eq!(package.body_table_hidden_axes(0usize)?, expected);
        Ok(())
    }

    #[test]
    fn concurrent_reads_and_edits_are_independent_copy_on_write_snapshots() -> TestResult {
        let source = normal_package()?;
        let package = Package::from_bytes(&source)?;
        let original = package.exact_bytes();

        // A regular `Package::clone` is a private COW fork: publishing an
        // edit from it must not mutate the source snapshot or its source
        // bytes.  The threaded checks below then exercise the same property
        // through `Arc<Package>` sharing.
        let fork = package.clone();
        let fork_target = HiddenAxes::new([AxisIndex::column(0)])?;
        let fork_before = fork.exact_bytes();
        let fork_result = fork
            .edit_body_table_hidden_axes(1usize)?
            .set(fork_target)
            .commit();
        assert!(matches!(
            fork_result,
            Err(Error::UnsupportedDependency | Error::UnsupportedSource)
        ));
        assert_eq!(fork.exact_bytes(), fork_before);
        assert_eq!(package.exact_bytes(), original);
        assert_eq!(package.body_table_hidden_axes(1usize)?, HiddenAxes::empty());

        let package = Arc::new(package);
        let start = Arc::new(Barrier::new(WORKERS));
        let mut handles = Vec::with_capacity(WORKERS);

        for worker in 0..WORKERS {
            let package = Arc::clone(&package);
            let start = Arc::clone(&start);
            handles.push(thread::spawn(move || {
                start.wait();
                if worker % 2 == 0 {
                    for _ in 0..READ_ROUNDS {
                        assert_eq!(
                            package
                                .body_table_hidden_axes(0usize)
                                .expect("reader sees the source snapshot"),
                            HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])
                                .expect("fixture axes are valid")
                        );
                    }
                } else {
                    let target = if worker % 4 == 1 {
                        HiddenAxes::new([AxisIndex::row(0), AxisIndex::column(3)])
                            .expect("row/column edit is valid")
                    } else {
                        HiddenAxes::new([AxisIndex::row(3)]).expect("row edit is valid")
                    };
                    let result = package
                        .edit_body_table_hidden_axes(1usize)
                        .expect("editor opens from an immutable source")
                        .set(target.clone())
                        .commit();
                    assert!(matches!(
                        result,
                        Err(Error::UnsupportedDependency | Error::UnsupportedSource)
                    ));
                    assert_eq!(
                        package
                            .body_table_hidden_axes(1usize)
                            .expect("source remains unchanged"),
                        HiddenAxes::empty()
                    );
                }
            }));
        }

        for handle in handles {
            handle.join().expect("read/edit worker must not panic");
        }
        assert_eq!(package.exact_bytes(), original);
        assert_eq!(
            package.body_table_hidden_axes(0usize)?,
            HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?
        );
        assert_eq!(package.body_table_hidden_axes(1usize)?, HiddenAxes::empty());
        Ok(())
    }

    #[test]
    fn one_patch_can_be_shared_across_concurrent_apply_and_inverse_cycles() -> TestResult {
        let source = normal_package()?;
        let package = Package::from_bytes(&source)?;
        let before = package.exact_bytes();
        let requested =
            HiddenAxes::new([AxisIndex::row(0), AxisIndex::row(3), AxisIndex::column(1)])?;
        let changed = package
            .edit_body_table_hidden_axes(0usize)?
            .set(requested.clone())
            .commit()?;
        let target = changed.package().exact_bytes();
        let patch = Arc::new(changed.patch().clone());
        let source = Arc::new(package);
        let start = Arc::new(Barrier::new(WORKERS));
        let mut handles = Vec::with_capacity(WORKERS);

        for _ in 0..WORKERS {
            let source = Arc::clone(&source);
            let patch = Arc::clone(&patch);
            let before = before.clone();
            let target = target.clone();
            let requested = requested.clone();
            let start = Arc::clone(&start);
            handles.push(thread::spawn(move || {
                start.wait();
                for _ in 0..READ_ROUNDS {
                    let applied = source
                        .apply_body_table_hidden_axes(&patch)
                        .expect("shared patch applies to its exact source");
                    assert_eq!(applied.package().exact_bytes(), target);
                    assert_eq!(
                        applied
                            .package()
                            .body_table_hidden_axes(0usize)
                            .expect("applied snapshot can be read"),
                        requested
                    );

                    let restored = applied
                        .package()
                        .apply_body_table_hidden_axes(&patch.inverse())
                        .expect("shared inverse applies to each private target");
                    assert_eq!(restored.package().exact_bytes(), before);
                    assert_eq!(
                        restored
                            .package()
                            .body_table_hidden_axes(0usize)
                            .expect("restored snapshot can be read"),
                        patch.before().clone()
                    );
                }
            }));
        }

        for handle in handles {
            handle.join().expect("patch worker must not panic");
        }
        assert_eq!(source.exact_bytes(), before);
        assert_eq!(patch.after(), &requested);
        Ok(())
    }

    #[test]
    fn independently_opened_snapshots_have_identical_first_access() -> TestResult {
        let source = Arc::new(normal_package()?);
        let expected = HiddenAxes::new([AxisIndex::row(1), AxisIndex::column(2)])?;
        let start = Arc::new(Barrier::new(WORKERS));
        let mut handles = Vec::with_capacity(WORKERS);

        for _ in 0..WORKERS {
            let source = Arc::clone(&source);
            let expected = expected.clone();
            let start = Arc::clone(&start);
            handles.push(thread::spawn(move || {
                start.wait();
                let package = Package::from_bytes(source.as_slice()).expect("source reopens");
                for _ in 0..READ_ROUNDS {
                    assert_eq!(
                        package
                            .body_table_hidden_axes(0usize)
                            .expect("independent first access succeeds"),
                        expected
                    );
                }
                assert_eq!(package.exact_bytes(), source.as_slice());
            }));
        }

        for handle in handles {
            handle.join().expect("independent reader must not panic");
        }
        Ok(())
    }
}
