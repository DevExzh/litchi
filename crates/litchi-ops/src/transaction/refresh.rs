//! [`File::refresh`] and [`Edit::refresh`]: one field refresh is one
//! mutation and at most one native commit (record 0279).

use std::sync::Arc;

use super::{Commit, Diagnostics, Edit, Error, History, LimitKind, Patch, SemanticLimits};
use crate::refresh::{Options, Summary};
use crate::{Code, File, Result};

impl File {
    /// Recompute the derived results the document caches, in one atomic
    /// commit: DOCX tables of contents and computable fields, PPTX slide
    /// numbers and dates, XLSX formula results (with
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
    /// [`Code::LimitExceeded`] past the default [`SemanticLimits`],
    /// [`Code::DocumentProtected`] for a signed or locked file, and
    /// [`Code::UnsupportedCapability`] for other formats (and for formula
    /// cache write-back outside XLSX).
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
    /// assert_eq!(report.value.count(Status::Stale), 1);
    /// assert!(report.patch.is_noop());
    /// let written = file.refresh(&Options::new().write_cache(true))?;
    /// assert_eq!(written.value.refreshed(), 1);
    /// let again = written.file.refresh(&Options::new().write_cache(true))?;
    /// assert_eq!(again.value.count(Status::Unchanged), 1);
    /// assert!(again.patch.is_noop());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn refresh(&self, options: &Options) -> Result<Commit<Summary>> {
        self.refresh_with_limits(options, &SemanticLimits::default())
    }

    /// [`Self::refresh`] under caller-selected transaction limits.
    ///
    /// # Errors
    /// As [`Self::refresh`].
    pub fn refresh_with_limits(
        &self,
        options: &Options,
        limits: &SemanticLimits,
    ) -> Result<Commit<Summary>> {
        let result = (|| -> Result<Commit<Summary>> {
            #[cfg(feature = "docx")]
            let admitted = matches!(self, File::Docx(_));
            #[cfg(not(feature = "docx"))]
            let admitted = false;
            if !admitted {
                super::admit(self)?;
            }
            let (step, summary) = run(self, options, limits)?;
            let mut steps = Vec::new();
            let mut snapshot_bytes = 0;
            if let Some(step) = step {
                if step.snapshot_bytes > limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: step.snapshot_bytes,
                        maximum: limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                snapshot_bytes = step.snapshot_bytes;
                steps.push(step);
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
            Ok(Commit {
                file,
                patch,
                diagnostics,
                value: summary,
            })
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
        let result = (|| {
            if index >= self.limits.max_operations() {
                return Err(Error::Limit {
                    kind: LimitKind::Operations,
                    observed: index.saturating_add(1),
                    maximum: self.limits.max_operations(),
                }
                .into());
            }
            let (step, summary) = run(&self.current, options, &self.limits)?;
            if let Some(step) = step {
                let bytes = self.snapshot_bytes.saturating_add(step.snapshot_bytes);
                if bytes > self.limits.max_snapshot_bytes() {
                    return Err(Error::Limit {
                        kind: LimitKind::SnapshotBytes,
                        observed: bytes,
                        maximum: self.limits.max_snapshot_bytes(),
                    }
                    .into());
                }
                self.steps
                    .try_reserve(1)
                    .map_err(|source| Error::Allocation { source })?;
                self.current = step.after.clone();
                self.steps.push(step);
                self.snapshot_bytes = bytes;
            }
            self.operations += 1;
            Ok(summary)
        })();
        result.map_err(|error| self.poison(error, index))
    }
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
        #[cfg(all(feature = "eval", any(feature = "xlsb", feature = "xls")))]
        _ if matches!(
            file.format(),
            Some(litchi_core::FileFormat::Xlsb | litchi_core::FileFormat::Xls)
        ) =>
        {
            if options.writes_cache() {
                return Err(crate::Error::new(
                    Code::UnsupportedCapability,
                    "formula cache write-back is implemented for XLSX workbooks",
                )
                .with_suggestion("refresh without write_cache for a report, or convert to XLSX"));
            }
            let report = file.recalculate(options.calc_options())?;
            Ok((
                None,
                crate::refresh::workbook_summary(file, &report, false)?,
            ))
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
        "field refresh recomputes DOCX fields, PPTX slide fields and workbook formulas",
    )
    .with_suggestion("refresh a DOCX, PPTX or XLSX file");
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}
