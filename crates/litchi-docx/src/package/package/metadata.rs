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
        let (package, changed, _) =
            self.publish_metadata(patch.target(), patch.parts(), patch.core())?;
        Ok(crate::metadata::Commit::new(
            package,
            patch.clone(),
            litchi_core::edit::Diagnostics::new(changed),
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

    /// Publish `target` as the core properties of a new package, and `parts`
    /// over its application and custom properties. Returns whether the
    /// package changed.
    pub(crate) fn publish_metadata(
        &self,
        target: Option<&Props>,
        parts: &families::Patch,
        retained_core: Option<&families::Patch>,
    ) -> Result<(Self, bool, Option<families::Patch>)> {
        let core = self.metadata() != target;
        let parts_noop = parts.is_noop();
        // A no-op patch from the shared properties planner carries a complete
        // property-family scope. Authenticate it before a core edit changes
        // the package; applying it after that edit would compare the scope
        // against the new core bytes and reject a valid combined edit.
        if parts_noop {
            parts.check(&self.opc)?;
        }
        if !core && parts_noop {
            return Ok((self.clone(), false));
        }
        if core && let Some(props) = target {
            admit(props)?;
        }
        if self.opc.is_signed() || self.opc.requires_signature_edit_policy() {
            return Err(Error::Signed);
        }
        let mut package = self.clone();
        let mut failure = None;
        let mut published_core = None;
        let published = package.edit_semantic_opc("edit_metadata", |opc| {
            if core {
                let core_commit = if let Some(retained_core) = retained_core {
                    retained_core.apply(opc).map_err(crate::Error::from)?
                } else {
                    let mut core_edit =
                        families::Edit::new(opc, Host::Word).map_err(crate::Error::from)?;
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
                    core_edit.commit().map_err(crate::Error::from)?
                };
                published_core = Some(core_commit.patch().clone());
                *opc = core_commit.into_package();
            }
            if !parts_noop {
                if let Err(error) = parts.apply(opc).map(|commit| *opc = commit.into_package()) {
                    failure = Some(error);
                    return Err(crate::Error::InvalidFormat(
                        "document-property parts refused the patch".to_owned(),
                    ));
                }
            }
            Ok(())
        });
        if let Some(error) = failure {
            return Err(Error::Properties(error));
        }
        published?;
        Ok((package, true, published_core))
    }
}
