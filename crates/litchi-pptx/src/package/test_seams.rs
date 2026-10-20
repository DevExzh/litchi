//! Test fixture seams: the families of record 0351 applied in place, so the
//! crate's source-level tests keep their fixtures compact. Every seam runs
//! the public family route (edit, verb, commit) and replaces the package
//! with the committed one; an error leaves the package unchanged.

use litchi_opc::PackURI;

use super::Package;
use crate::Result;

impl Package {
    fn replace_with(&mut self, package: Self) {
        *self = package;
    }

    pub(crate) fn add_slide_master(&mut self) -> Result<crate::master_layout::AuthoredSlideMaster> {
        let mut edit = self.edit_master_layout()?;
        let value = edit.add_slide_master()?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(value)
    }

    pub(crate) fn add_slide_layout(
        &mut self,
        master_part_name: &PackURI,
        kind: crate::master_layout::SlideLayoutKind,
        name: &str,
        placeholders: &[crate::master_layout::PlaceholderSpec],
    ) -> Result<crate::master_layout::AuthoredSlideLayout> {
        let mut edit = self.edit_master_layout()?;
        let value = edit.add_slide_layout(master_part_name, kind, name, placeholders)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(value)
    }

    pub(crate) fn remove_slide_layout(&mut self, layout_part_name: &PackURI) -> Result<()> {
        let mut edit = self.edit_master_layout()?;
        edit.remove_slide_layout(layout_part_name)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(())
    }

    pub(crate) fn store_placeholder_shape(
        &mut self,
        part_name: &PackURI,
        spec: &crate::master_layout::PlaceholderSpec,
    ) -> Result<()> {
        let mut edit = self.edit_master_layout()?;
        edit.set_placeholder_shape(part_name, spec)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(())
    }

    pub(crate) fn put_custom_props(&mut self, props: crate::custom::Props) -> Result<()> {
        let metadata = |error: crate::metadata::Error| match error {
            crate::metadata::Error::Signed => {
                crate::Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
            },
            error => crate::Error::Invalid(error.to_string()),
        };
        let mut edit = self.edit_metadata().map_err(metadata)?;
        edit.set_custom_props(props);
        let package = edit.commit().map_err(metadata)?.into_package();
        self.replace_with(package);
        Ok(())
    }

    pub(crate) fn put_math_properties(
        &mut self,
        value: crate::presentation_properties::math::Properties,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        let mut edit = self.edit_math_properties()?;
        let previous = edit.set(value)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(previous)
    }

    pub(crate) fn remove_math_properties(
        &mut self,
    ) -> Result<Option<crate::presentation_properties::math::Properties>> {
        let mut edit = self.edit_math_properties()?;
        let previous = edit.remove()?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(previous)
    }

    pub(crate) fn put_zooms<'a>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'a>>,
        owner: crate::shape::zoom::Owner,
    ) -> Result<Option<crate::shape::zoom::Owner>> {
        let mut edit = self.edit_zooms()?;
        let previous = edit.set(slide, owner)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(previous)
    }

    pub(crate) fn put_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
        value: bool,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        let mut edit = self.edit_shape_design_element()?;
        let previous = edit.set(slide, shape, value)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(previous)
    }

    pub(crate) fn remove_shape_design_element<'s, 'k>(
        &mut self,
        slide: impl Into<crate::slide::Selector<'s>>,
        shape: impl Into<crate::shape::Key<'k>>,
    ) -> Result<Option<crate::shape::designer::Snapshot>> {
        let mut edit = self.edit_shape_design_element()?;
        let previous = edit.remove(slide, shape)?;
        let package = edit.commit()?.into_package();
        self.replace_with(package);
        Ok(previous)
    }

    pub(crate) fn apply_placeholder_type_extension_slot_patch(
        &mut self,
        change: &crate::master_layout::placeholder::SlotChange,
    ) -> Result<crate::master_layout::placeholder::Snapshot> {
        let mut staged = crate::part_edit::Staged::new(self, "apply_placeholder_type_extension")?;
        let snapshot = staged
            .update(|opc| crate::master_layout::placeholder::apply_slot_patch(opc, change))?;
        let (package, _) = staged.commit()?;
        self.replace_with(package);
        Ok(snapshot)
    }
}
