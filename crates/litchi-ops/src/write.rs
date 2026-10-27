//! Immutable operation and publication dispatch.

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "ods",
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
use std::io;
use std::io::Write;

use crate::path::IntoPath;
use crate::target::IntoTarget;
use crate::{Added, Code, Error, File, IntoNode, IntoPlacement, IntoProps, Result};

impl File {
    /// Add an element and return its new snapshot and canonical path.
    ///
    /// Placement defaults to the end of the parent. Properties are admitted
    /// before the isolated leaf transaction is committed.
    /// Legacy PPT inserts blank slides (`layout` may be `"blank"`) under the
    /// presentation root, at any placement.
    /// XLS adds an empty worksheet tab under the workbook (optional `name`), or
    /// inserts one empty row or column under a sheet at an explicit position.
    /// PPTX slide insertion requires an exact `layout` property and an eager
    /// native owner (created presentations or `SourceMode::Eager` opens). The
    /// inserted slide is empty and shares the selected layout.
    /// Add a `shape` (or `textbox`) under a slide to append a text box. Positive `width` and
    /// `height` are required; `x` and `y` default to zero and `text` to empty.
    /// Lengths accept the shared unit syntax. Text-box placement is end-only.
    ///
    /// # Errors
    /// Returns target, property, protection, source or capability errors.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_ops::{NewNode, open};
    /// let file = open("report.docx")?;
    /// let added = file.add("/body", NewNode::new("paragraph").prop("text", "Hello"))?;
    /// let paragraph = added.file.get(&added.path)?;
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn add<'p, 'n>(
        &self,
        _parent: impl IntoPath<'p>,
        _node: impl IntoNode<'n>,
    ) -> Result<Added> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => {
                let (handle, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Docx(handle.into_file()),
                    path,
                })
            },
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Xlsx(commit.into_file()),
                    path,
                })
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Pptx(commit.into_file()),
                    path,
                })
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Xlsb(commit.into_file()),
                    path,
                })
            },
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Rtf(commit.after),
                    path,
                })
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: commit.after.into_file(),
                    path,
                })
            },

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: commit.after.into_file(),
                    path,
                })
            },
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: commit.after.into_file(),
                    path,
                })
            },
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Doc(commit.after),
                    path,
                })
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Ppt(commit.after),
                    path,
                })
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                let (commit, path) = handle.add_with_limits(
                    _parent.into_path()?,
                    _node.into_node()?,
                    &crate::transaction::SemanticLimits::default(),
                )?;
                Ok(Added {
                    file: Self::Xls(commit.after),
                    path,
                })
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Ods(commit.after),
                    path,
                })
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => {
                let (commit, path) = handle.add(_parent.into_path()?, _node.into_node()?)?;
                Ok(Added {
                    file: Self::Odp(commit.after),
                    path,
                })
            },
            Self::Signed(_) => Err(protected(self)),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "adding elements")),
        }
    }

    /// Move an element to a sibling position or another supported parent.
    ///
    /// Legacy PPT supports slide ordering and XLS complete-catalog tab
    /// ordering through their native package edits. ODS moves
    /// dependency-free sheets and rows through its leaf splices.
    /// Positional destinations are measured after detaching the element.
    /// The original snapshot stays unchanged on success and failure.
    ///
    /// # Errors
    /// Returns target, placement, protection, source or capability errors.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_ops::{Placement, open};
    /// let file = open("report.docx")?;
    /// let moved = file.move_to("/body/p[1]", Placement::End)?;
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn move_to<'p, 'q>(
        &self,
        _path: impl IntoPath<'p>,
        _to: impl IntoPlacement<'q>,
    ) -> Result<File> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Docx(commit.into_file())),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Xlsx(commit.into_file())),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Xlsb(commit.into_file())),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Ppt(commit.after)),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle
                .move_to_with_limits(
                    _path.into_path()?,
                    _to.into_placement()?,
                    &crate::transaction::SemanticLimits::default(),
                )
                .map(|commit| Self::Xls(commit.after)),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Pptx(commit.into_file())),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Rtf(commit.after)),
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| commit.after.into_file()),

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Doc(commit.after)),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Ods(commit.after)),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle
                .move_to(_path.into_path()?, _to.into_placement()?)
                .map(|commit| Self::Odp(commit.after)),
            Self::Signed(_) => Err(protected(self)),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "moving elements")),
        }
    }

    /// Remove one path or every match of a scoped query in one transaction.
    ///
    /// The original snapshot remains unchanged. Selection and dependency
    /// checks refuse removals that the format cannot preserve safely.
    /// PPTX supports dependency-free slides in created or eagerly opened
    /// presentations and requires at least one slide to remain.
    /// XLSB removes complete-catalog sheet tabs, or stored cell records from
    /// cells, ranges, and cell queries; virtual cells are exact no-ops and
    /// neighboring cells never shift.  A sheet query must be homogeneous, and
    /// complete-tab removal retains at least one visible tab.
    /// Legacy PPT supports slide removal subject to native dependency admission.
    /// XLS removes one worksheet tab (dependency-checked; a visible tab must
    /// remain), one empty row or column, or stored cell records from cells,
    /// ranges, and cell queries; virtual cells are exact no-ops and worksheet
    /// neighbors never shift.
    ///
    /// # Errors
    /// Returns target, protection, source, dependency or capability errors.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// let file = litchi_ops::open("report.docx")?;
    /// let updated = file.remove("/body/p[1]/r[1]")?;
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn remove<'p>(&self, _target: impl IntoTarget<'p>) -> Result<File> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Docx(commit.into_file())),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Xlsx(commit.into_file())),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Xlsb(commit.into_file())),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Xls(commit.after)),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Ppt(commit.after)),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Pptx(commit.into_file())),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Rtf(commit.after)),
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| commit.after.into_file()),

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Doc(commit.after)),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Ods(commit.after)),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle
                .remove(_target.into_target()?)
                .map(|commit| Self::Odp(commit.after)),
            Self::Signed(_) => Err(protected(self)),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "removing elements")),
        }
    }

    /// Exchange two elements while preserving their content and formatting.
    ///
    /// PPTX supports slides in created or eagerly opened presentations.
    /// Legacy PPT supports slide swaps on retained native owners.
    /// XLSB and XLS support complete-catalog tab ordering, including chart
    /// tabs, in retained source and eager owners.
    /// RTF supports ordinary main-story paragraph swaps through its native
    /// lifecycle editor.
    /// Numbers swaps two sheets, or two tables of one sheet or of two
    /// sheets, as the two moves applied one at a time; its rows and columns
    /// are refused (record 0347).
    ///
    /// Both paths resolve against this snapshot. The returned handle owns the
    /// complete change; a failure leaves the original unchanged.
    ///
    /// # Errors
    /// Returns target, protection, source, dependency or capability errors.
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// let file = litchi_ops::open("report.docx")?;
    /// let updated = file.swap("/body/p[1]", "/body/p[3]")?;
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn swap<'p, 'q>(&self, _a: impl IntoPath<'p>, _b: impl IntoPath<'q>) -> Result<File> {
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Docx(commit.into_file())),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Xlsx(commit.into_file())),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Xlsb(commit.into_file())),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Ppt(commit.after)),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle
                .swap_with_limits(
                    _a.into_path()?,
                    _b.into_path()?,
                    &crate::transaction::SemanticLimits::default(),
                )
                .map(|commit| Self::Xls(commit.after)),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Pptx(commit.into_file())),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Rtf(commit.after)),
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| commit.after.into_file()),

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Doc(commit.after)),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Ods(commit.after)),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle
                .swap(_a.into_path()?, _b.into_path()?)
                .map(|commit| Self::Odp(commit.after)),
            Self::Signed(_) => Err(protected(self)),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "swapping elements")),
        }
    }

    /// Set properties on a path or a scoped query, returning a new snapshot.
    ///
    /// The original handle and its clones remain unchanged on success and
    /// failure. All assignments belong to one isolated leaf edit. Available
    /// properties and format support are described by [`Self::capabilities`].
    ///
    /// XLS and XLSX cells, rectangle ranges and scoped cell selectors share one
    /// native commit. Ranges include blank coordinates. XLS admits scalar values,
    /// clears and sheet names; formula cells require a separate write handler.
    /// XLSX formulas broadcast literally. Its sheet and row/column layout
    /// assignments also use one native commit per call, including bounded
    /// selectors of one target family.
    /// PPTX slide shapes and placeholders accept text replacement or clearing;
    /// scoped selectors share one native commit.
    /// Legacy PPT top-level shapes and placeholders accept text and geometry;
    /// slides accept `hidden`. PPT anchors round to the nearest 1/576-inch
    /// master unit. Native closure checks may refuse unsupported records.
    /// Default [`crate::transaction::SemanticLimits`] bound expanded targets
    /// and staged payloads. Use [`Self::edit_with_options`] to select those bounds.
    ///
    /// A `find` property turns the call into a text search over the target's
    /// subtree, as OfficeCLI's `set --find` (record 0273): `replace` gives
    /// the replacement (empty deletes), `matchCase` (default `true`) and
    /// `wholeWord` (default `false`) tune matching, and every other property
    /// formats the replaced text (or, without `replace`, the matched text).
    /// A match spanning runs takes the formatting of the run where it
    /// starts; untouched runs keep their bytes. Zero matches are an exact
    /// no-op. DOCX, PPTX, XLSX and RTF support it ([`crate::capabilities::Verb::Replace`]);
    /// [`Self::apply`] reports the count as [`crate::Applied::matched`].
    ///
    /// ```
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_core::FileFormat;
    /// use litchi_ops::{NewNode, create, props};
    /// let file = create(FileFormat::Docx)?;
    /// let file = file
    ///     .add("/body", NewNode::new("paragraph").prop("text", "draft one, draft two"))?
    ///     .file;
    /// let file = file.set("/", props! { "find" => "draft", "replace" => "final" })?;
    /// let text = file.get("/body/p[1]")?.view().text()?.unwrap_or_default().into_owned();
    /// assert_eq!(text, "final one, final two");
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    ///
    /// # Errors
    /// Refuses unknown or invalid properties, unsupported targets, unscoped
    /// mutation queries, protected documents and stale source snapshots. A
    /// find refuses a regular expression (`regex=true`), an empty pattern,
    /// and a call that neither replaces nor formats (`missing_property`).
    ///
    /// ```no_run
    /// # #[cfg(feature = "docx")]
    /// # {
    /// use litchi_ops::{open, props};
    /// let original = open("report.docx")?;
    /// let updated = original.set("/body/p[1]", props! { "text" => "Revised text" })?;
    /// assert_eq!(updated.format(), original.format());
    /// # }
    /// # Ok::<(), litchi_ops::Error>(())
    /// ```
    pub fn set<'p, 'v>(
        &self,
        _target: impl IntoTarget<'p>,
        _props: impl IntoProps<'v>,
    ) -> Result<File> {
        if let Self::Signed(signed) = self {
            // A signed setter may still prove an exact no-op through the leaf
            // editor. Keep the outer wrapper and its original publication;
            // a changed native commit remains protected and can never strip
            // the signature implicitly.
            let target = _target.into_target()?;
            let props = crate::Props::from_props(_props)?;
            let updated = signed.file().set(target, props)?;
            if !crate::transaction::same_content_for_replay(&updated, signed.file(), true)? {
                return Err(protected(self));
            }
            return Ok(self.clone());
        }
        let _props = match self {
            Self::Locked(locked) => return Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => crate::Props::from_props(_props)?,
        };
        if crate::find::requested(&_props) {
            return crate::find::set(self, _target.into_target()?, &_props);
        }
        match self {
            #[cfg(feature = "docx")]
            Self::Docx(handle) => handle
                .set_file(&_target.into_target()?, _props)
                .map(Self::Docx),
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Xlsb(commit.into_file())),
            #[cfg(feature = "xls")]
            Self::Xls(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Xls(commit.after)),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Ppt(commit.after)),
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Xlsx(commit.into_file())),
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Pptx(commit.into_file())),
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Doc(commit.after)),
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Rtf(commit.after)),
            #[cfg(feature = "ods")]
            Self::Ods(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Ods(commit.after)),
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| commit.after.into_file()),

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| commit.after.into_file()),
            #[cfg(feature = "odp")]
            Self::Odp(handle) => handle
                .set(_target.into_target()?, _props)
                .map(|commit| Self::Odp(commit.after)),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "setting properties")),
        }
    }

    /// Publish this snapshot into an owned byte buffer in its original format.
    ///
    /// # Errors
    /// Returns source, output-policy and format capability errors. Publication
    /// is repeatable and does not consume or mutate this handle.
    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        match self {
            // A retained DOC or PPT package publishes its exact compound-file
            // bytes; the leaf copies them once into a buffer of the known
            // size, where the generic sink route below would copy every
            // chunk a second time (record 0377).
            #[cfg(feature = "doc")]
            Self::Doc(handle) => handle.to_bytes(),
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => handle.to_bytes(),
            _ => {
                let mut bytes = Vec::new();
                self.write_to(&mut bytes)?;
                Ok(bytes)
            },
        }
    }

    /// Publish this snapshot to a caller-owned sink and return bytes written.
    ///
    /// Existing signatures are preserved; an explicitly unsigned handle
    /// publishes signature removal through the leaf. The sink may contain a
    /// prefix if it fails; publication never modifies the input snapshot.
    ///
    /// # Errors
    /// Returns typed leaf, source, encryption-policy and I/O failures, or
    /// `unsupported_capability` for formats whose publisher is not integrated.
    #[cfg_attr(
        not(any(
            feature = "doc",
            feature = "docx",
            feature = "ppt",
            feature = "xlsx",
            feature = "pptx",
            feature = "xlsb",
            feature = "xls",
            feature = "rtf",
            feature = "ods",
            feature = "odt",
            feature = "odp",
            odf_objects,
            feature = "iwork-adapter"
        )),
        allow(
            clippy::only_used_in_recursion,
            reason = "signed publication forwards the sink when no publisher feature is enabled"
        )
    )]
    pub fn write_to<W: Write>(&self, sink: W) -> Result<u64> {
        match self {
            #[cfg(feature = "rtf")]
            Self::Rtf(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "odt")]
            Self::Odt(handle) | Self::Oth(handle) | Self::Odm(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },

            #[cfg(odf_objects)]
            Self::Odg(handle)
            | Self::Odc(handle)
            | Self::Odi(handle)
            | Self::Odb(handle)
            | Self::Odf(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "iwork-adapter")]
            Self::Pages(handle) | Self::Keynote(handle) | Self::Numbers(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "xls")]
            Self::Xls(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "docx")]
            Self::Docx(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "ppt")]
            Self::Ppt(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "xlsb")]
            Self::Xlsb(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "xlsx")]
            Self::Xlsx(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "pptx")]
            Self::Pptx(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "doc")]
            Self::Doc(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "ods")]
            Self::Ods(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            #[cfg(feature = "odp")]
            Self::Odp(handle) => {
                let mut counted = Counted { sink, bytes: 0 };
                handle.write_to(&mut counted)?;
                Ok(counted.bytes)
            },
            Self::Signed(signed) => signed.file().write_to(sink),
            Self::Locked(locked) => Err(locked.read_error()),
            #[allow(unreachable_patterns)]
            _ => Err(pending(self, "publishing")),
        }
    }
}

fn protected(file: &File) -> Error {
    formatted(
        file,
        Error::new(
            Code::DocumentProtected,
            "a signed document cannot be changed",
        ),
    )
    .with_suggestion("consume the signature with `Signed::unsign` before editing")
}

fn pending(file: &File, operation: &str) -> Error {
    formatted(
        file,
        Error::new(
            Code::UnsupportedCapability,
            format!("{operation} this format is not integrated yet"),
        ),
    )
    .with_suggestion("see the format capability report and M2 follow-up")
}

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "ods",
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
pub(crate) struct Counted<W> {
    pub(crate) sink: W,
    pub(crate) bytes: u64,
}

#[cfg(any(
    feature = "doc",
    feature = "docx",
    feature = "ppt",
    feature = "xlsx",
    feature = "pptx",
    feature = "xlsb",
    feature = "xls",
    feature = "rtf",
    feature = "ods",
    feature = "odt",
    feature = "odp",
    odf_objects,
    feature = "iwork-adapter"
))]
impl<W: Write> Write for Counted<W> {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = self.sink.write(bytes)?;
        self.bytes = self
            .bytes
            .checked_add(
                u64::try_from(count)
                    .map_err(|_error| io::Error::other("output byte count exceeds u64"))?,
            )
            .ok_or_else(|| io::Error::other("output byte count exceeds u64"))?;
        Ok(count)
    }
    fn flush(&mut self) -> io::Result<()> {
        self.sink.flush()
    }
}

fn formatted(file: &File, error: Error) -> Error {
    match file.format() {
        Some(format) => error.with_format(format),
        None => error,
    }
}
