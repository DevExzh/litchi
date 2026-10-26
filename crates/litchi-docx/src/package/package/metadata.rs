//! Package entry points of the core-properties transaction family.

use litchi_ooxml_common::custom::{Host, Props as CustomProps};
use litchi_ooxml_common::properties::{self as families, app};

use super::super::model::{Package, Props};
use crate::metadata::{Edit, Error, Patch, Result, admit};

impl Package {
    /// Borrow the typed core properties, retaining package absence: `None`
    /// means the package has no core-properties part.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// assert!(package.metadata().is_some());
    /// # Ok::<(), litchi_docx::Error>(())
    /// ```
    #[must_use]
    pub fn metadata(&self) -> Option<&Props> {
        self.properties.get()
    }

    /// Stage a core-properties change over this snapshot.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_metadata();
    /// edit.update(|props| props.subject = Some("Plans".to_owned()));
    /// let next = edit.commit()?.into_package();
    /// assert_eq!(
    ///     next.metadata().and_then(|props| props.subject.as_deref()),
    ///     Some("Plans")
    /// );
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    #[must_use]
    pub fn edit_metadata(&self) -> Edit {
        Edit::new(self.clone(), self.metadata().cloned())
    }

    /// Replay a metadata patch onto this snapshot and return its commit.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_metadata();
    /// edit.update(|props| props.title = Some("Draft".to_owned()));
    /// let (changed, patch) = edit.commit()?.into_parts();
    /// let restored = changed.apply_metadata(&patch.inverse())?.into_package();
    /// assert_eq!(restored.metadata(), package.metadata());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package's core properties
    /// differ from the patch's base, and the metadata publication errors
    /// otherwise.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<crate::metadata::Commit> {
        if self.metadata() != patch.base() {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.ensure_opc_current("apply_metadata")?;
            patch.whole().check(&self.opc).map_err(Error::Properties)?;
            return Ok(crate::metadata::Commit::new(
                self.clone(),
                patch.clone(),
                litchi_core::edit::Diagnostics::new(false),
            ));
        }
        let mut package = self.clone();
        package.edit_semantic_opc("apply_metadata", |opc| {
            let commit = patch.whole().apply(opc).map_err(crate::Error::from)?;
            *opc = commit.into_package();
            Ok(())
        })?;
        Ok(crate::metadata::Commit::new(
            package,
            patch.clone(),
            litchi_core::edit::Diagnostics::new(true),
        ))
    }

    /// The application properties (`docProps/app.xml`), when present.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// assert!(package.app_props()?.is_some());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::Properties`] for an invalid part or graph.
    pub fn app_props(&self) -> Result<Option<app::Props>> {
        app::read(&self.opc).map_err(|error| Error::Properties(error.into()))
    }

    /// The exact part patch that stages `app` and `custom` (each `None`:
    /// untouched) over this package's application and custom properties
    /// (record 0342).
    pub(crate) fn property_parts_patch(
        &self,
        app: Option<Option<app::Props>>,
        custom: Option<CustomProps>,
        remove_custom: bool,
    ) -> Result<families::Patch> {
        // A dirty cached value would make the part bytes stale.
        self.ensure_opc_current("edit_metadata")?;
        let snapshot = families::Snapshot::read(&self.opc, Host::Word)?;
        let mut plan = snapshot.plan();
        if let Some(app) = app {
            plan.set_app_props(app);
        }
        if let Some(custom) = custom {
            if remove_custom {
                plan.remove_custom_props();
            } else {
                plan.set_custom_props(custom);
            }
        }
        Ok(plan.finish()?.1)
    }

    /// Publish the staged core, application and custom properties as one
    /// atomic package edit. The core family is planned first; the application
    /// and custom families are then planned against that intermediate graph so
    /// package relationship IDs cannot collide. The returned private patch is
    /// the composition of those exact family patches.
    pub(crate) fn publish_metadata(
        &self,
        target: Option<&Props>,
        app: Option<Option<app::Props>>,
        custom: Option<CustomProps>,
        remove_custom: bool,
    ) -> Result<(Self, bool, families::Patch, families::Patch)> {
        let core = self.metadata() != target;
        let has_parts = app.is_some() || custom.is_some() || remove_custom;
        if !core && !has_parts {
            return Ok((
                self.clone(),
                false,
                families::Patch::default(),
                families::Patch::default(),
            ));
        }
        if core && let Some(props) = target {
            admit(props)?;
        }

        // An application/custom-only edit can use the existing source plan.
        // This both preserves signed no-op behavior and avoids planning the
        // same graph twice. A core edit must defer this plan until after the
        // core family has been applied to the candidate.
        if !core {
            let parts = if has_parts {
                self.property_parts_patch(app.clone(), custom.clone(), remove_custom)?
            } else {
                families::Patch::default()
            };
            if parts.is_noop() {
                return Ok((self.clone(), false, parts.clone(), parts));
            }
            if self.opc.is_signed() || self.opc.requires_signature_edit_policy() {
                return Err(Error::Signed);
            }
            let mut package = self.clone();
            let mut parts_failure = None;
            let published =
                package.edit_semantic_opc("edit_metadata", |opc| match parts.apply(opc) {
                    Ok(commit) => {
                        *opc = commit.into_package();
                        Ok(())
                    },
                    Err(error) => {
                        parts_failure = Some(error);
                        Err(crate::Error::InvalidFormat(
                            "document-property parts refused the patch".to_owned(),
                        ))
                    },
                });
            if let Some(error) = parts_failure {
                return Err(Error::Properties(error));
            }
            published?;
            return Ok((package, true, parts.clone(), parts));
        }

        if self.opc.is_signed() || self.opc.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }

        let mut package = self.clone();
        let mut parts_failure = None;
        let mut compose_failure = None;
        let mut public_parts = families::Patch::default();
        let mut whole = families::Patch::default();
        let published = package.edit_semantic_opc("edit_metadata", |opc| {
            let mut core_edit = families::Edit::new(opc, Host::Word).map_err(crate::Error::from)?;
            match target {
                Some(props) => {
                    core_edit
                        .set_core_metadata(props.clone())
                        .map_err(crate::Error::from)?;
                },
                None => {
                    core_edit.clear_metadata();
                },
            }
            let core_commit = core_edit.commit().map_err(crate::Error::from)?;
            let (core_package, core_patch) = core_commit.into_parts();
            *opc = core_package;

            let parts_patch = if has_parts {
                let mut parts_edit =
                    families::Edit::new(opc, Host::Word).map_err(crate::Error::from)?;
                if let Some(app) = app {
                    parts_edit.set_app_props(app);
                }
                if let Some(custom) = custom {
                    if remove_custom {
                        parts_edit.remove_custom_props();
                    } else {
                        parts_edit.set_custom_props(custom);
                    }
                }
                match parts_edit.commit() {
                    Ok(commit) => {
                        let (parts_package, parts_patch) = commit.into_parts();
                        *opc = parts_package;
                        Some(parts_patch)
                    },
                    Err(error) => {
                        parts_failure = Some(error);
                        return Err(crate::Error::InvalidFormat(
                            "document-property parts refused the patch".to_owned(),
                        ));
                    },
                }
            } else {
                None
            };
            public_parts = parts_patch.clone().unwrap_or_default();
            whole = if let Some(parts_patch) = parts_patch {
                match core_patch.compose(&parts_patch) {
                    Ok(composed) => composed,
                    Err(error) => {
                        compose_failure = Some(error);
                        return Err(crate::Error::InvalidFormat(
                            "document-property family patches could not be composed".to_owned(),
                        ));
                    },
                }
            } else {
                core_patch
            };
            Ok(())
        });
        if let Some(error) = parts_failure {
            return Err(Error::Properties(error));
        }
        if let Some(error) = compose_failure {
            return Err(Error::Properties(error));
        }
        published?;
        Ok((package, !whole.is_noop(), public_parts, whole))
    }
}
