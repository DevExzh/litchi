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

    /// Replay a metadata patch onto this snapshot and return a new package.
    ///
    /// ```
    /// use litchi_docx::Package;
    ///
    /// let package = Package::new()?;
    /// let mut edit = package.edit_metadata();
    /// edit.update(|props| props.title = Some("Draft".to_owned()));
    /// let (changed, patch) = edit.commit()?.into_parts();
    /// let restored = changed.apply_metadata(&patch.inverse())?;
    /// assert_eq!(restored.metadata(), package.metadata());
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package's core properties
    /// differ from the patch's base, and the [`Edit::commit`] errors
    /// otherwise.
    pub fn apply_metadata(&self, patch: &Patch) -> Result<Self> {
        if self.metadata() != patch.base() {
            return Err(Error::PatchConflict);
        }
        self.publish_metadata(patch.target(), patch.parts())
            .map(|(package, _)| package)
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
    ) -> Result<families::Patch> {
        // A dirty cached value would make the part bytes stale.
        self.ensure_opc_current("edit_metadata")?;
        let snapshot = families::Snapshot::read(&self.opc, Host::Word)?;
        let mut plan = snapshot.plan();
        if let Some(app) = app {
            plan.set_app_props(app);
        }
        if let Some(custom) = custom {
            plan.set_custom_props(custom);
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
    ) -> Result<(Self, bool)> {
        let core = self.metadata() != target;
        if !core && parts.is_noop() {
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
        let published = package.edit_semantic_opc("edit_metadata", |opc| {
            if core {
                match target {
                    Some(props) => litchi_ooxml_common::properties::write(opc, props)?,
                    None => litchi_ooxml_common::properties::clear(opc)?,
                };
            }
            if let Err(error) = parts.apply(opc) {
                failure = Some(error);
                return Err(crate::Error::InvalidFormat(
                    "document-property parts refused the patch".to_owned(),
                ));
            }
            Ok(())
        });
        if let Some(error) = failure {
            return Err(Error::Properties(error));
        }
        published?;
        Ok((package, true))
    }
}
