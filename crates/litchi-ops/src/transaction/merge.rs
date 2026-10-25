//! [`File::merge`] and [`Edit::merge`]: one template merge is one mutation
//! and at most one native commit (record 0274).

use std::sync::Arc;

use super::{
    Commit, Diagnostics, Edit, EditOptions, Error, History, LimitKind, Patch, SemanticLimits,
};
use crate::merge::{Options, Record, Summary};
use crate::{Code, File, Result};

impl File {
    /// Fill `record` into this template's `{{key}}` placeholders and, for
    /// DOCX, its `MERGEFIELD` fields, in one atomic commit. See
    /// [`crate::merge`] for the placeholder, field, missing-key and limit
    /// rules. The template is unchanged, so one handle fills any number of
    /// records.
    ///
    /// Supported for DOCX, PPTX, XLSX, XLSB, DOC, PPT, XLS, RTF, ODT, OTH,
    /// ODM, ODS and ODP (records 0274 and 0295). A template without a matching
    /// placeholder or field is an exact no-op: the returned file is the
    /// template and its patch is a no-op.
    ///
    /// # Errors
    ///
    /// [`Code::InvalidValue`] for a value a document cannot store,
    /// [`Code::LimitExceeded`] past the default [`EditOptions`],
    /// [`Code::DocumentProtected`] for a signed or locked file, and
    /// [`Code::UnsupportedCapability`] for other formats. The template
    /// handle is unchanged on failure.
    ///
    /// ```
    /// # #[cfg(feature = "xlsx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::merge::{Options, Record};
    /// use litchi_ops::props;
    ///
    /// let template = litchi_ops::create(FileFormat::Xlsx)?
    ///     .set("/Sheet1/A1", props! { "value" => "Invoice for {{client}}" })?;
    /// let record: Record = [("client", "Acme")].into_iter().collect();
    /// let commit = template.merge(&record, &Options::new())?;
    /// let cell = commit.file().get("/Sheet1/A1")?;
    /// assert_eq!(cell.view().text()?.as_deref(), Some("Invoice for Acme"));
    /// assert_eq!(commit.value().replacements(), 1);
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn merge(&self, record: &Record, options: &Options) -> Result<Commit<Summary>> {
        self.merge_with_options(record, options, EditOptions::default())
    }

    /// [`Self::merge`] under caller-selected [`EditOptions`]: raise the
    /// semantic limits' `max_expanded_targets` and `max_staged_bytes` for
    /// large templates.
    ///
    /// # Errors
    /// As [`Self::merge`].
    pub fn merge_with_options(
        &self,
        record: &Record,
        options: &Options,
        edit_options: EditOptions,
    ) -> Result<Commit<Summary>> {
        let limits = &edit_options.semantic_limits_value();
        // One merge is the transaction `File::edit` would build around
        // `Edit::merge`, without its separate admission editor: a DOCX merge
        // builds its editor first, which applies the same signature, source
        // and protection policy, so a second one would only repeat it.
        let result = (|| -> Result<Commit<Summary>> {
            #[cfg(feature = "docx")]
            let admitted = matches!(self, File::Docx(_));
            #[cfg(not(feature = "docx"))]
            let admitted = false;
            if !admitted {
                super::admit(self)?;
            }
            let (steps, summary) = run_steps(self, record, options, limits)?;
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
    /// Fill `record` into the current candidate as one mutation of this
    /// transaction; see [`File::merge`].
    ///
    /// # Errors
    /// As [`File::merge`], with the mutation's item index.
    pub fn merge(&mut self, record: &Record, options: &Options) -> Result<Summary> {
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
            let (steps, summary) = run_steps(&self.current, record, options, &operation_limits)?;
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

/// Merge through the format's native editor: the native commits, each on
/// the previous one's result (one, except for chunked XLS edits).
fn run_steps(
    file: &File,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Vec<super::Step>, Summary)> {
    #[cfg(feature = "xls")]
    if let File::Xls(handle) = file {
        let (commits, summary) =
            crate::format::xls::write::find::merge(handle, record, options, limits)?;
        return Ok((super::xls_steps(commits)?, summary));
    }
    let (step, summary) = run(file, record, options, limits)?;
    Ok((step.into_iter().collect(), summary))
}

/// Merge through the format's native editor: at most one native commit.
fn run(
    file: &File,
    record: &Record,
    options: &Options,
    limits: &SemanticLimits,
) -> Result<(Option<super::Step>, Summary)> {
    match file {
        #[cfg(feature = "docx")]
        File::Docx(handle) => {
            let (commit, summary) =
                crate::format::docx::merge::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "pptx")]
        File::Pptx(handle) => {
            let (commit, summary) =
                crate::format::pptx::write::merge::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_pptx(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "xlsx")]
        File::Xlsx(handle) => {
            let (commit, summary) =
                crate::format::xlsx::write::merge::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_xlsx(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "rtf")]
        File::Rtf(handle) => {
            let (commit, summary) = handle.merge(record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_rtf(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "doc")]
        File::Doc(handle) => {
            let (commit, summary) = handle.merge(record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_doc(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "xlsb")]
        File::Xlsb(handle) => {
            let (commit, summary) =
                crate::format::xlsb::write::find::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_xlsb(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "ods")]
        File::Ods(handle) => {
            let (commit, summary) =
                crate::format::ods::write::find::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_ods(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "odt")]
        File::Odt(handle) | File::Oth(handle) | File::Odm(handle) => {
            let (commit, summary) =
                crate::format::odt::write::find::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_odt(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "odp")]
        File::Odp(handle) => {
            let (commit, summary) =
                crate::format::odp::write::find::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_odp(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        #[cfg(feature = "ppt")]
        File::Ppt(handle) => {
            let (commit, summary) =
                crate::format::ppt::write::find::merge(handle, record, options, limits)?;
            match commit {
                Some(commit) => Ok((Some(super::wrap_ppt(handle, commit, ())?.0), summary)),
                None => Ok((None, summary)),
            }
        },
        _ => {
            let _ = (record, options, limits);
            Err(unsupported(file))
        },
    }
}

fn unsupported(file: &File) -> crate::Error {
    let error = crate::Error::new(
        Code::UnsupportedCapability,
        "template merge fills word-processing, presentation and spreadsheet templates",
    )
    .with_suggestion(
        "convert the template to a word-processing, presentation or spreadsheet format",
    );
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}
