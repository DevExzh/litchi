//! Test fixture helpers: the edit families of record 0351 applied in place.
//!
//! Each helper runs the public family route (`edit_x`, a verb, `commit`) and
//! replaces the fixture package with the committed one; an error leaves the
//! package unchanged. They keep fixtures compact in tests whose subject is
//! not the family itself.

#![allow(
    dead_code,
    reason = "each integration test uses a subset of the helpers"
)]

use litchi_opc::{OpcPackage, PackURI};
use litchi_pptx::{Package, Result};

/// Replace `package` with the package `edit` commits.
macro_rules! in_place {
    ($package:ident, $edit:ident, |$e:ident| $verb:expr) => {{
        let mut $e = $package.$edit()?;
        let value = $verb?;
        let committed = $e.commit()?.into_package();
        *$package = committed;
        Ok(value)
    }};
}

/// The record 0351 families applied in place.
pub(crate) trait InPlace {
    fn edit_opc_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T>;
    fn put_tags_in_place<'a>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'a>>,
        list: litchi_pptx::tag::List,
    ) -> Result<Option<litchi_pptx::tag::List>>;
    fn remove_tags_in_place<'a>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'a>>,
    ) -> Result<Option<litchi_pptx::tag::List>>;
    fn put_shape_tags_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        list: litchi_pptx::tag::List,
    ) -> Result<Option<litchi_pptx::tag::List>>;
    fn remove_shape_tags_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::tag::List>>;
    fn put_styles_in_place(&mut self, styles: litchi_pptx::table::style::List) -> Result<bool>;
    fn remove_styles_in_place(&mut self) -> Result<Option<litchi_pptx::table::style::List>>;
    fn put_fonts_in_place(&mut self, fonts: litchi_pptx::font::Fonts) -> Result<bool>;
    fn remove_fonts_in_place(&mut self) -> Result<Option<litchi_pptx::font::Fonts>>;
    fn put_model3d_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        model: litchi_pptx::model3d::Model,
    ) -> Result<Option<litchi_pptx::model3d::Model>>;
    fn remove_model3d_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::model3d::Model>>;
    fn put_shape_classification_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        outcome: litchi_pptx::shape::classification::Outcome,
    ) -> Result<Option<litchi_pptx::shape::classification::Snapshot>>;
    fn remove_shape_classification_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::shape::classification::Snapshot>>;
    fn put_shape_design_element_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        value: bool,
    ) -> Result<Option<litchi_pptx::shape::designer::Snapshot>>;
    fn remove_shape_design_element_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::shape::designer::Snapshot>>;
    fn put_shape_designer_properties_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        properties: litchi_pptx::shape::designer::DrawingProperties,
    ) -> Result<litchi_pptx::shape::designer::properties::Snapshot>;
    fn remove_shape_designer_properties_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<litchi_pptx::shape::designer::properties::Snapshot>;
    fn put_slide_designer_tags_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        tags: litchi_pptx::shape::designer::Tags,
    ) -> Result<Option<litchi_pptx::shape::designer::Tags>>;
    fn remove_slide_designer_tags_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<Option<litchi_pptx::shape::designer::Tags>>;
    fn put_zooms_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        owner: litchi_pptx::shape::zoom::Owner,
    ) -> Result<Option<litchi_pptx::shape::zoom::Owner>>;
    fn remove_zooms_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<Option<litchi_pptx::shape::zoom::Owner>>;
    fn put_math_properties_in_place(
        &mut self,
        value: litchi_pptx::presentation_properties::math::Properties,
    ) -> Result<Option<litchi_pptx::presentation_properties::math::Properties>>;
    fn remove_math_properties_in_place(
        &mut self,
    ) -> Result<Option<litchi_pptx::presentation_properties::math::Properties>>;
    fn put_read_only_recommended_in_place(&mut self, value: bool) -> Result<Option<bool>>;
    fn remove_read_only_recommended_in_place(&mut self) -> Result<Option<bool>>;
    fn put_custom_props_in_place(&mut self, props: litchi_pptx::custom::Props) -> Result<()>;
    fn put_notes_in_place(&mut self, graph: litchi_pptx::notes::Graph) -> Result<bool>;
    fn remove_notes_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<bool>;
    fn clear_notes_in_place(&mut self) -> Result<usize>;
    fn add_slide_master_in_place(&mut self) -> Result<litchi_pptx::AuthoredSlideMaster>;
    fn add_slide_layout_in_place(
        &mut self,
        master_part_name: &PackURI,
        kind: litchi_pptx::SlideLayoutKind,
        name: &str,
        placeholders: &[litchi_pptx::PlaceholderSpec],
    ) -> Result<litchi_pptx::AuthoredSlideLayout>;
    fn remove_slide_layout_in_place(&mut self, layout_part_name: &PackURI) -> Result<()>;
    fn store_placeholder_shape_in_place(
        &mut self,
        part_name: &PackURI,
        spec: &litchi_pptx::PlaceholderSpec,
    ) -> Result<()>;
    fn stage_cross_slide_copy_in_place(
        &mut self,
        source: &Package,
        plan: &litchi_pptx::cross_slide_copy::Plan,
    ) -> Result<()>;
    fn apply_cross_slide_copy_in_place(
        &mut self,
        source: &Package,
        patch: &litchi_pptx::cross_slide_copy::Patch,
    ) -> Result<()>;
    fn edit_guides_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::guides::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_structure_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::structure::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_custom_shows_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::custom_show::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_changes_info_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::changes::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_show_events_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::events::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_slide_ole_objects_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation::embedded::ole::slide::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_comments_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::comments::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_modern_comments_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::modern_comments::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_themes_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::shape::theme::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_ink_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::ink::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_media_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::media_parts::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_charts_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::chart::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_laser_traces_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::laser::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_ole_objects_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::ole::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
    fn edit_revision_info_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::revision::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    fn edit_slide_sync_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::slide_sync::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T>;
    #[cfg(feature = "vba-inspection")]
    fn edit_vba_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::vba::Edit<'_>) -> Result<T>,
    ) -> Result<T>;
}

impl InPlace for Package {
    fn edit_opc_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut OpcPackage) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_opc, |edit| edit.update(operation))
    }

    fn put_tags_in_place<'a>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'a>>,
        list: litchi_pptx::tag::List,
    ) -> Result<Option<litchi_pptx::tag::List>> {
        let package = self;
        in_place!(package, edit_tags, |edit| edit.set(slide, list))
    }

    fn remove_tags_in_place<'a>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'a>>,
    ) -> Result<Option<litchi_pptx::tag::List>> {
        let package = self;
        in_place!(package, edit_tags, |edit| edit.remove(slide))
    }

    fn put_shape_tags_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        list: litchi_pptx::tag::List,
    ) -> Result<Option<litchi_pptx::tag::List>> {
        let package = self;
        in_place!(package, edit_shape_tags, |edit| edit
            .set(slide, shape, list))
    }

    fn remove_shape_tags_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::tag::List>> {
        let package = self;
        in_place!(package, edit_shape_tags, |edit| edit.remove(slide, shape))
    }

    fn put_styles_in_place(&mut self, styles: litchi_pptx::table::style::List) -> Result<bool> {
        let package = self;
        in_place!(package, edit_styles, |edit| edit.set(styles))
    }

    fn remove_styles_in_place(&mut self) -> Result<Option<litchi_pptx::table::style::List>> {
        let package = self;
        in_place!(package, edit_styles, |edit| edit.remove())
    }

    fn put_fonts_in_place(&mut self, fonts: litchi_pptx::font::Fonts) -> Result<bool> {
        let package = self;
        in_place!(package, edit_fonts, |edit| edit.set(fonts))
    }

    fn remove_fonts_in_place(&mut self) -> Result<Option<litchi_pptx::font::Fonts>> {
        let package = self;
        in_place!(package, edit_fonts, |edit| edit.remove())
    }

    fn put_model3d_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        model: litchi_pptx::model3d::Model,
    ) -> Result<Option<litchi_pptx::model3d::Model>> {
        let package = self;
        in_place!(package, edit_model3d, |edit| edit.set(slide, shape, model))
    }

    fn remove_model3d_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::model3d::Model>> {
        let package = self;
        in_place!(package, edit_model3d, |edit| edit.remove(slide, shape))
    }

    fn put_shape_classification_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        outcome: litchi_pptx::shape::classification::Outcome,
    ) -> Result<Option<litchi_pptx::shape::classification::Snapshot>> {
        let package = self;
        in_place!(package, edit_shape_classification, |edit| edit
            .set(slide, shape, outcome))
    }

    fn remove_shape_classification_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::shape::classification::Snapshot>> {
        let package = self;
        in_place!(package, edit_shape_classification, |edit| edit
            .remove(slide, shape))
    }

    fn put_shape_design_element_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        value: bool,
    ) -> Result<Option<litchi_pptx::shape::designer::Snapshot>> {
        let package = self;
        in_place!(package, edit_shape_design_element, |edit| edit
            .set(slide, shape, value))
    }

    fn remove_shape_design_element_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<Option<litchi_pptx::shape::designer::Snapshot>> {
        let package = self;
        in_place!(package, edit_shape_design_element, |edit| edit
            .remove(slide, shape))
    }

    fn put_shape_designer_properties_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
        properties: litchi_pptx::shape::designer::DrawingProperties,
    ) -> Result<litchi_pptx::shape::designer::properties::Snapshot> {
        let package = self;
        in_place!(package, edit_shape_designer_properties, |edit| edit
            .set(slide, shape, properties))
    }

    fn remove_shape_designer_properties_in_place<'s, 'k>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        shape: impl Into<litchi_pptx::shape::Key<'k>>,
    ) -> Result<litchi_pptx::shape::designer::properties::Snapshot> {
        let package = self;
        in_place!(package, edit_shape_designer_properties, |edit| edit
            .remove(slide, shape))
    }

    fn put_slide_designer_tags_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        tags: litchi_pptx::shape::designer::Tags,
    ) -> Result<Option<litchi_pptx::shape::designer::Tags>> {
        let package = self;
        in_place!(package, edit_slide_designer_tags, |edit| edit
            .set(slide, tags))
    }

    fn remove_slide_designer_tags_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<Option<litchi_pptx::shape::designer::Tags>> {
        let package = self;
        in_place!(package, edit_slide_designer_tags, |edit| edit.remove(slide))
    }

    fn put_zooms_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
        owner: litchi_pptx::shape::zoom::Owner,
    ) -> Result<Option<litchi_pptx::shape::zoom::Owner>> {
        let package = self;
        in_place!(package, edit_zooms, |edit| edit.set(slide, owner))
    }

    fn remove_zooms_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<Option<litchi_pptx::shape::zoom::Owner>> {
        let package = self;
        in_place!(package, edit_zooms, |edit| edit.remove(slide))
    }

    fn put_math_properties_in_place(
        &mut self,
        value: litchi_pptx::presentation_properties::math::Properties,
    ) -> Result<Option<litchi_pptx::presentation_properties::math::Properties>> {
        let package = self;
        in_place!(package, edit_math_properties, |edit| edit.set(value))
    }

    fn remove_math_properties_in_place(
        &mut self,
    ) -> Result<Option<litchi_pptx::presentation_properties::math::Properties>> {
        let package = self;
        in_place!(package, edit_math_properties, |edit| edit.remove())
    }

    fn put_read_only_recommended_in_place(&mut self, value: bool) -> Result<Option<bool>> {
        let package = self;
        in_place!(package, edit_read_only_recommended, |edit| edit.set(value))
    }

    fn remove_read_only_recommended_in_place(&mut self) -> Result<Option<bool>> {
        let package = self;
        in_place!(package, edit_read_only_recommended, |edit| edit.remove())
    }

    fn put_custom_props_in_place(&mut self, props: litchi_pptx::custom::Props) -> Result<()> {
        let metadata = |error: litchi_pptx::metadata::Error| match error {
            litchi_pptx::metadata::Error::Signed => {
                litchi_pptx::Error::Opc(litchi_opc::OpcError::SignedSourceRequiresExplicitPolicy)
            },
            error => litchi_pptx::Error::Invalid(error.to_string()),
        };
        let mut edit = self.edit_metadata().map_err(metadata)?;
        edit.set_custom_props(props);
        let committed = edit.commit().map_err(metadata)?.into_package();
        *self = committed;
        Ok(())
    }

    fn put_notes_in_place(&mut self, graph: litchi_pptx::notes::Graph) -> Result<bool> {
        let package = self;
        in_place!(package, edit_notes, |edit| edit.set(graph))
    }

    fn remove_notes_in_place<'s>(
        &mut self,
        slide: impl Into<litchi_pptx::slide::Selector<'s>>,
    ) -> Result<bool> {
        let package = self;
        in_place!(package, edit_notes, |edit| edit.remove(slide))
    }

    fn clear_notes_in_place(&mut self) -> Result<usize> {
        let package = self;
        in_place!(package, edit_notes, |edit| edit.clear())
    }

    fn add_slide_master_in_place(&mut self) -> Result<litchi_pptx::AuthoredSlideMaster> {
        let package = self;
        in_place!(package, edit_master_layout, |edit| edit.add_slide_master())
    }

    fn add_slide_layout_in_place(
        &mut self,
        master_part_name: &PackURI,
        kind: litchi_pptx::SlideLayoutKind,
        name: &str,
        placeholders: &[litchi_pptx::PlaceholderSpec],
    ) -> Result<litchi_pptx::AuthoredSlideLayout> {
        let master = master_id(self, master_part_name)?;
        let package = self;
        in_place!(package, edit_master_layout, |edit| edit.add_slide_layout(
            master,
            kind,
            name,
            placeholders
        ))
    }

    fn remove_slide_layout_in_place(&mut self, layout_part_name: &PackURI) -> Result<()> {
        let layout = layout_id(self, layout_part_name)?;
        let package = self;
        in_place!(package, edit_master_layout, |edit| edit
            .remove_slide_layout(layout))
    }

    fn store_placeholder_shape_in_place(
        &mut self,
        part_name: &PackURI,
        spec: &litchi_pptx::PlaceholderSpec,
    ) -> Result<()> {
        let owner = owner(self, part_name)?;
        let package = self;
        in_place!(package, edit_master_layout, |edit| edit
            .set_placeholder_shape(owner, spec))
    }

    fn stage_cross_slide_copy_in_place(
        &mut self,
        source: &Package,
        plan: &litchi_pptx::cross_slide_copy::Plan,
    ) -> Result<()> {
        let package = self;
        in_place!(package, edit, |edit| edit
            .stage_cross_slide_copy(source, plan))
    }

    fn apply_cross_slide_copy_in_place(
        &mut self,
        source: &Package,
        patch: &litchi_pptx::cross_slide_copy::Patch,
    ) -> Result<()> {
        let committed = self.apply_cross_slide_copy(source, patch)?.into_package();
        *self = committed;
        Ok(())
    }

    fn edit_guides_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::guides::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_guides, |edit| operation(&mut edit))
    }

    fn edit_structure_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::structure::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_structure, |edit| operation(&mut edit))
    }

    fn edit_custom_shows_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::custom_show::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_custom_shows, |edit| operation(&mut edit))
    }

    fn edit_changes_info_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::changes::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_changes_info, |edit| operation(&mut edit))
    }

    fn edit_show_events_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::events::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_show_events, |edit| operation(&mut edit))
    }

    fn edit_slide_ole_objects_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation::embedded::ole::slide::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_slide_ole_objects, |edit| operation(&mut edit))
    }
    fn edit_comments_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::comments::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_comments, |edit| operation(&mut edit))
    }

    fn edit_modern_comments_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::modern_comments::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_modern_comments, |edit| operation(&mut edit))
    }

    fn edit_themes_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::shape::theme::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_themes, |edit| operation(&mut edit))
    }

    fn edit_ink_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::ink::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_ink, |edit| operation(&mut edit))
    }

    fn edit_media_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::media_parts::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_media, |edit| operation(&mut edit))
    }

    fn edit_charts_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::chart::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_charts, |edit| operation(&mut edit))
    }

    fn edit_laser_traces_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::laser::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_laser_traces, |edit| operation(&mut edit))
    }

    fn edit_ole_objects_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::ole::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_ole_objects, |edit| operation(&mut edit))
    }

    fn edit_revision_info_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::revision::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_revision_info, |edit| operation(&mut edit))
    }

    fn edit_slide_sync_in_place<T>(
        &mut self,
        operation: impl FnOnce(
            &mut litchi_pptx::presentation_properties::metadata::slide_sync::Edit<'_>,
        ) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_slide_sync, |edit| operation(&mut edit))
    }

    #[cfg(feature = "vba-inspection")]
    fn edit_vba_in_place<T>(
        &mut self,
        operation: impl FnOnce(&mut litchi_pptx::presentation::embedded::vba::Edit<'_>) -> Result<T>,
    ) -> Result<T> {
        let package = self;
        in_place!(package, edit_vba, |edit| operation(&mut edit))
    }
}

/// The typed ID of the slide master whose part is `part_name`.
pub(crate) fn master_id(
    package: &Package,
    part_name: &PackURI,
) -> Result<litchi_pptx::slide::MasterId> {
    for master in package.presentation()?.slide_masters()? {
        if master.part().part().partname() == part_name {
            return master.id();
        }
    }
    Err(litchi_pptx::Error::Invalid(format!(
        "no slide master {part_name}"
    )))
}

/// The typed ID of the slide layout whose part is `part_name`.
pub(crate) fn layout_id(
    package: &Package,
    part_name: &PackURI,
) -> Result<litchi_pptx::slide::LayoutId> {
    for layout in package.presentation()?.slide_layouts()? {
        if layout.part().part().partname() == part_name {
            return layout.id();
        }
    }
    Err(litchi_pptx::Error::Invalid(format!(
        "no slide layout {part_name}"
    )))
}

/// The slide, layout or master whose part is `part_name`.
pub(crate) fn owner(
    package: &Package,
    part_name: &PackURI,
) -> Result<litchi_pptx::slide::Owner<'static>> {
    if let Ok(id) = master_id(package, part_name) {
        return Ok(id.into());
    }
    if let Ok(id) = layout_id(package, part_name) {
        return Ok(id.into());
    }
    let presentation = package.presentation()?;
    for index in 0..presentation.slide_count()? {
        if let Some(slide) = presentation.slide(index)?
            && slide.part().part().partname() == part_name
        {
            return Ok(index.into());
        }
    }
    Err(litchi_pptx::Error::Invalid(format!(
        "no owner part {part_name}"
    )))
}
