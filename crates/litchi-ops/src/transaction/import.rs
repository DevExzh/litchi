//! [`File::import`] and [`Edit::import`]: one CSV/TSV import is one
//! mutation and one native commit (record 0266).

use super::{Commit, Edit, EditOptions, Error, LimitKind, SemanticLimits};
use crate::import::{self, Options, Summary};
use crate::path::IntoPath;
use crate::{Code, File, Result};

impl File {
    /// Import CSV or TSV text into a worksheet at `target`: a sheet
    /// (`/Sheet1`, origin `A1`) or a cell (`/Sheet1/B2`, the origin), in one
    /// atomic commit. See [`crate::import`] for the text, type and limit
    /// rules.
    ///
    /// Supported for XLSX, XLSB and XLS. Empty input is an exact no-op that
    /// still validates the target.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidInput`] for malformed text, [`Code::InvalidArgument`]
    /// for unusable options, [`Code::InvalidValue`] for an over-long field,
    /// [`Code::OutOfRange`] when the data leaves the grid,
    /// [`Code::LimitExceeded`] past the default
    /// [`EditOptions`], path errors for the target, and
    /// [`Code::UnsupportedCapability`] for other formats. The original handle
    /// is unchanged on failure.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::import::Options;
    ///
    /// let file = litchi_ops::create(FileFormat::Xlsx)?;
    /// let commit = file.import("/Sheet1/B2", b"1,2\n3,4\n", &Options::csv())?;
    /// assert_eq!((commit.value().rows(), commit.value().columns()), (2, 2));
    /// let cell = commit.file().get("/Sheet1/C3")?;
    /// assert_eq!(cell.view().text()?.as_deref(), Some("4"));
    /// assert!(file.import("/Sheet1", b"", &Options::csv())?.patch().is_noop());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn import<'p>(
        &self,
        target: impl IntoPath<'p>,
        source: &[u8],
        options: &Options,
    ) -> Result<Commit<Summary>> {
        self.import_with_options(target, source, options, EditOptions::default())
    }

    /// [`Self::import`] under caller-selected [`EditOptions`]: raise the
    /// semantic limits' `max_expanded_targets` and `max_staged_bytes` for
    /// large sources.
    ///
    /// # Errors
    /// As [`Self::import`].
    pub fn import_with_options<'p>(
        &self,
        target: impl IntoPath<'p>,
        source: &[u8],
        options: &Options,
        edit_options: EditOptions,
    ) -> Result<Commit<Summary>> {
        self.edit_with_options(edit_options, |edit| edit.import(target, source, options))
    }
}

impl Edit {
    /// Import CSV or TSV text as one mutation of this transaction; see
    /// [`File::import`].
    ///
    /// # Errors
    /// As [`File::import`], with the mutation's item index.
    pub fn import<'p>(
        &mut self,
        target: impl IntoPath<'p>,
        source: &[u8],
        options: &Options,
    ) -> Result<Summary> {
        self.check()?;
        self.flush_pending()?;
        let index = self.operations;
        let operation_limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        let result = (|| {
            let path = target.into_path()?;
            let table = import::parse(source, options, &self.limits)?;
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let summary = if table.is_empty() {
                Summary::at(origin(&self.current, &path)?, &table)
            } else {
                let (steps, summary) = run(&self.current, &path, &table, &operation_limits)?;
                self.record(steps)?;
                summary
            };
            self.operations += 1;
            Ok(summary)
        })();
        result.map_err(|error| self.poison(error, index))
    }
}

impl Edit {
    /// Retain the native steps of one import, as `execute` retains one
    /// step, charging their snapshots to the transaction.
    fn record(&mut self, steps: Vec<super::Step>) -> Result<()> {
        let mut bytes = self.snapshot_bytes;
        for step in &steps {
            bytes = bytes.saturating_add(step.snapshot_bytes);
        }
        if bytes > self.limits.max_snapshot_bytes() {
            return Err(Error::Limit {
                kind: LimitKind::SnapshotBytes,
                observed: bytes,
                maximum: self.limits.max_snapshot_bytes(),
            }
            .into());
        }
        self.steps
            .try_reserve(steps.len())
            .map_err(|source| Error::Allocation { source })?;
        for step in steps {
            self.current = step.after.clone();
            self.steps.push(step);
        }
        self.snapshot_bytes = bytes;
        Ok(())
    }
}

/// Write a non-empty table through the format's native editor: one native
/// commit, except that XLS commits every chunk its staged-change bound
/// admits.
fn run(
    file: &File,
    path: &crate::path::PathArg<'_>,
    table: &import::Table,
    limits: &SemanticLimits,
) -> Result<(Vec<super::Step>, Summary)> {
    match file {
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => {
            let (commit, summary) =
                crate::format::xlsx::write::import::import(handle, path, table, limits)?;
            let (step, summary) = super::wrap_xlsx(handle, commit, summary)?;
            Ok((vec![step], summary))
        },
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => {
            let (commit, summary) =
                crate::format::xlsb::write::import::import(handle, path, table, limits)?;
            let (step, summary) = super::wrap_xlsb(handle, commit, summary)?;
            Ok((vec![step], summary))
        },
        #[cfg(feature = "xls")]
        File::Xls(handle) => {
            let (commits, summary) =
                crate::format::xls::write::import::import(handle, path, table, limits)?;
            let mut steps = Vec::new();
            for commit in commits {
                let before = commit.before.clone();
                steps.push(super::wrap_xls(&before, commit, ())?.0);
            }
            Ok((steps, summary))
        },
        #[cfg(feature = "ods")]
        File::Ods(handle) => {
            let (commit, summary) =
                crate::format::ods::write::import::import(handle, path, table, limits)?;
            let (step, summary) = super::wrap_ods(handle, commit, summary)?;
            Ok((vec![step], summary))
        },
        _ => {
            let _ = (path, table, limits);
            Err(unsupported(file))
        },
    }
}

/// Resolve the target of an empty import, which writes nothing.
fn origin(file: &File, path: &crate::path::PathArg<'_>) -> Result<import::Origin> {
    match file {
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => crate::format::xlsx::write::import::origin(handle, path),
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => crate::format::xlsb::write::import::origin(handle, path),
        #[cfg(feature = "xls")]
        File::Xls(handle) => crate::format::xls::write::import::origin(handle, path),
        #[cfg(feature = "ods")]
        File::Ods(handle) => crate::format::ods::write::import::origin(handle, path),
        _ => {
            let _ = path;
            Err(unsupported(file))
        },
    }
}

fn unsupported(file: &File) -> crate::Error {
    let error = crate::Error::new(
        Code::UnsupportedCapability,
        "import writes worksheets of XLSX, XLSB and XLS workbooks",
    );
    match file.format() {
        Some(litchi_core::FileFormat::Ods) => error
            .with_format(litchi_core::FileFormat::Ods)
            .with_suggestion("ODS has no unified write path yet (M6)"),
        Some(format) => error.with_format(format),
        None => error,
    }
}
