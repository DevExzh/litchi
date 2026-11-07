//! [`File::refresh`] and [`Edit::refresh`]: one field refresh is one
//! mutation and at most one native commit (record 0279).

use std::sync::Arc;

use super::{
    Commit, Diagnostics, Edit, EditOptions, Error, History, LimitKind, Patch, SemanticLimits,
};
use crate::refresh::{Options, Summary};
use crate::{Code, File, Result};

impl File {
    /// Recompute the derived results the document caches, in one atomic
    /// publication: DOC/DOCX computable fields, DOCX tables of contents,
    /// PPTX slide numbers and dates, and XLS/XLSX/XLSB formula results (with
    /// [`Options::write_cache`]). See [`crate::refresh`] for what each
    /// format computes, what it keeps, and how each field is reported.
    ///
    /// A document whose results are current is an exact no-op: the returned
    /// file is this one and its patch is a no-op. The handle is never
    /// changed.
    ///
    /// # Errors
    ///
    /// [`Code::RefreshFailed`] for a table of contents whose page numbers
    /// are unknown under [`PageNumbers::Require`](crate::refresh::PageNumbers::Require),
    /// [`Code::LimitExceeded`] past the default [`EditOptions`],
    /// [`Code::DocumentProtected`] for a signed or locked file, and
    /// [`Code::UnsupportedCapability`] for formats without a refresh adapter
    /// or a changed native text owner outside its edit profile (including
    /// PPTX markup-compatibility branch selection). XLS cache writes may
    /// use several native chunks inside this atomic unified operation.
    ///
    /// ```
    /// # #[cfg(all(feature = "xlsx", feature = "eval"))]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::props;
    /// use litchi_ops::refresh::{Options, Status};
    ///
    /// let file = litchi_ops::create(FileFormat::Xlsx)?
    ///     .set("/Sheet1/A1", props! { "value" => 20 })?
    ///     .set("/Sheet1/A2", props! { "formula" => "A1*2" })?;
    /// let report = file.refresh(&Options::new())?;
    /// assert_eq!(report.value().count(Status::Stale), 1);
    /// assert!(report.patch().is_noop());
    /// let written = file.refresh(&Options::new().write_cache(true))?;
    /// assert_eq!(written.value().refreshed(), 1);
    /// let again = written.file().refresh(&Options::new().write_cache(true))?;
    /// assert_eq!(again.value().count(Status::Unchanged), 1);
    /// assert!(again.patch().is_noop());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn refresh(&self, options: &Options) -> Result<Commit<Summary>> {
        self.refresh_with_options(options, EditOptions::default())
    }

    /// [`Self::refresh`] under caller-selected [`EditOptions`].
    ///
    /// # Errors
    /// As [`Self::refresh`].
    pub fn refresh_with_options(
        &self,
        options: &Options,
        edit_options: EditOptions,
    ) -> Result<Commit<Summary>> {
        let limits = &edit_options.semantic_limits_value();
        let result = (|| -> Result<Commit<Summary>> {
            #[cfg(feature = "docx")]
            let admitted = matches!(self, File::Docx(_));
            #[cfg(not(feature = "docx"))]
            let admitted = false;
            if !admitted {
                super::admit(self)?;
            }
            let (steps, summary) = run_steps(self, options, limits)?;
            let snapshot_bytes = steps.iter().fold(0usize, |bytes, step| {
                bytes.saturating_add(step.snapshot_bytes)
            });
            if snapshot_bytes > limits.max_snapshot_bytes() {
                return Err(Error::Limit {
                    kind: LimitKind::SnapshotBytes,
                    observed: snapshot_bytes,
                    maximum: limits.max_snapshot_bytes(),
                }
                .into());
            }
            let changed = super::changed(&steps)?;
            let file = match steps.last() {
                Some(step) if changed => step.after.clone(),
                _ => self.clone(),
            };
            let diagnostics = Diagnostics {
                changed,
                operations: 1,
                commits: steps.len(),
                leaf_operations: steps.iter().map(|step| step.leaf_operations).sum(),
                snapshot_bytes,
            };
            let patch = Patch {
                inner: Arc::new(History {
                    source: self.clone(),
                    target: file.clone(),
                    steps,
                    changed,
                }),
                inverse: false,
            };
            patch.check_sources()?;
            Ok(Commit::new(file, patch, diagnostics, summary))
        })()
        .map_err(|error: crate::Error| {
            if error.item().is_some() {
                error
            } else {
                error.with_item(0)
            }
        });
        super::check_source(self)?;
        result
    }
}

impl Edit {
    /// Refresh the current candidate as one mutation of this transaction;
    /// see [`File::refresh`].
    ///
    /// # Errors
    /// As [`File::refresh`], with the mutation's item index.
    pub fn refresh(&mut self, options: &Options) -> Result<Summary> {
        self.check()?;
        self.flush_pending()?;
        let index = self.operations;
        let operation_limits = self.limits.with_snapshot_budget(
            self.limits
                .max_snapshot_bytes()
                .saturating_sub(self.snapshot_bytes),
        );
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.checked_add(1).unwrap_or(usize::MAX),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let (steps, summary) = run_steps(&self.current, options, &operation_limits)?;
            let bytes = steps.iter().fold(self.snapshot_bytes, |bytes, step| {
                bytes.saturating_add(step.snapshot_bytes)
            });
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
            self.operations += 1;
            Ok(summary)
        })();
        result.map_err(|error| self.poison(error, index))
    }
}

/// Refresh through the format's native editor: the native commits, each
/// on the previous one's result (one, except for chunked XLS edits).
fn run_steps(
    file: &File,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Vec<super::Step>, Summary)> {
    #[cfg(all(feature = "xls", feature = "eval"))]
    if let File::Xls(handle) = file {
        let (commits, summary) =
            crate::format::xls::write::refresh::refresh(file, handle, options, limits)?;
        return Ok((super::xls_steps(commits)?, summary));
    }
    let (step, summary) = run(file, options, limits)?;
    Ok((step.into_iter().collect(), summary))
}

/// Refresh through the format's native editor: at most one native commit.
fn run(
    file: &File,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<super::Step>, Summary)> {
    match file {
        #[cfg(feature = "docx")]
        File::Docx(handle) => {
            let (commit, summary) = crate::format::docx::refresh::refresh(handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => {
            let (commit, summary) =
                crate::format::pptx::write::refresh::refresh(handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_pptx(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(all(feature = "xlsx", feature = "eval"))]
        File::Xlsx(handle) => {
            let (commit, summary) =
                crate::format::xlsx::write::refresh::refresh(file, handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_xlsx(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "doc")]
        File::Doc(handle) => {
            let (commit, summary) = crate::format::doc::refresh::refresh(handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_doc(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "odt")]
        File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => {
            let (commit, summary) =
                crate::format::odt::write::refresh::refresh(handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_odt(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "odp")]
        File::Odp(handle) => {
            let (commit, summary) =
                crate::format::odp::write::refresh::refresh(handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_odp(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(all(feature = "xlsb", feature = "eval"))]
        File::Xlsb(handle) => {
            let (commit, summary) =
                crate::format::xlsb::write::refresh::refresh(file, handle, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_xlsb(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        _ => {
            let _ = (options, limits);
            Err(unsupported(file))
        },
    }
}

fn unsupported(file: &File) -> crate::Error {
    let error = crate::Error::new(
        Code::UnsupportedCapability,
        "this format has no field refresh adapter",
    )
    .with_suggestion(
        "refresh a DOCX, DOC, ODT, PPTX, ODP or workbook file (record 0295 follow-up)",
    );
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}
