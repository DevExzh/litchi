//! Typed presentation authoring, finalized as a canonical package.
//!
//! [`Builder`] owns the authoring drafts — the [`Presentation`] with its
//! [`Slide`]s and their [`Shape`]s — and publishes them once, as an
//! immutable [`Package`], in [`Builder::build`] (record 0351 §2.1 rule 7).

pub(crate) mod presentation;
pub(crate) mod shape;
pub(crate) mod slide;

#[cfg(test)]
mod draft_tests;

pub use presentation::{FIRST_SLIDE_ID, Presentation};
pub use shape::{Shape, ShapeType};
pub use slide::Slide;

use crate::{Package, Result};

/// A native authoring owner whose output is an immutable-editable package.
///
/// Publication lives on the package returned by [`Self::build`].
///
/// ```
/// use litchi_pptx::builder::Builder;
/// let mut builder = Builder::new()?;
/// builder.presentation()?.add_slide()?
///     .add_text_box("Hello", 0, 0, 914_400, 914_400);
/// let package = builder.build()?;
/// let mut bytes = Vec::new();
/// package.write_to(&mut bytes)?;
/// # Ok::<(), litchi_pptx::Error>(())
/// ```
pub struct Builder {
    package: Package,
    presentation: Presentation,
    embed_fonts: bool,
    #[cfg(feature = "automatic-fonts")]
    font_preparation_limits: litchi_fonts::PreparationLimits,
    #[cfg(feature = "automatic-fonts")]
    font_eot_limits: litchi_fonts::embedding::powerpoint::Limits,
}
impl Builder {
    /// Start the native minimal empty presentation.
    ///
    /// # Errors
    /// Returns native construction failures.
    pub fn new() -> Result<Self> {
        Ok(Self {
            package: Package::new_authoring()?,
            presentation: Presentation::new(),
            embed_fonts: false,
            #[cfg(feature = "automatic-fonts")]
            font_preparation_limits: litchi_fonts::PreparationLimits::default(),
            #[cfg(feature = "automatic-fonts")]
            font_eot_limits: litchi_fonts::embedding::powerpoint::Limits::default(),
        })
    }

    /// Configure the native typed presentation before finalization.
    ///
    /// # Errors
    /// Returns native authoring admission failures.
    pub fn presentation(&mut self) -> Result<&mut Presentation> {
        Ok(&mut self.presentation)
    }

    /// Select automatic font embedding during native graph finalization.
    ///
    /// ```
    /// use litchi_pptx::builder::Builder;
    /// use litchi_fonts::embedding::Mode;
    /// let mut builder = Builder::new()?;
    /// builder.font_embedding(Mode::Full)?;
    /// // Add authored text before finalization to include its discovered fonts.
    /// let package = builder.build()?;
    /// assert_eq!(package.presentation()?.slide_count()?, 0);
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    /// Returns native font-policy admission failures. Font discovery and
    /// embedding failures are returned by [`Self::build`].
    #[cfg(feature = "automatic-fonts")]
    pub fn font_embedding(&mut self, mode: litchi_fonts::embedding::Mode) -> Result<&mut Self> {
        std::sync::Arc::make_mut(&mut self.package.opc).with_fonts(match mode {
            litchi_fonts::embedding::Mode::Full => litchi_opc::FontEmbedding::Full,
            litchi_fonts::embedding::Mode::Subset => litchi_opc::FontEmbedding::Subset,
        });
        self.embed_fonts = true;
        Ok(self)
    }

    /// Bound automatic font preparation during finalization.
    #[cfg(feature = "automatic-fonts")]
    #[must_use]
    pub fn font_preparation_limits(
        &mut self,
        limits: litchi_fonts::PreparationLimits,
    ) -> &mut Self {
        self.font_preparation_limits = limits;
        self
    }

    /// Bound PowerPoint EOT publication during finalization.
    #[cfg(feature = "automatic-fonts")]
    #[must_use]
    pub fn font_eot_limits(
        &mut self,
        limits: litchi_fonts::embedding::powerpoint::Limits,
    ) -> &mut Self {
        self.font_eot_limits = limits;
        self
    }

    /// Finalize the native graph without serializing and reopening a ZIP.
    ///
    /// # Errors
    /// Returns native graph construction and validation failures.
    pub fn build(self) -> Result<Package> {
        self.package.publish_authored(
            &self.presentation,
            self.embed_fonts,
            #[cfg(feature = "automatic-fonts")]
            self.font_preparation_limits,
            #[cfg(feature = "automatic-fonts")]
            self.font_eot_limits,
        )
    }
}
impl std::fmt::Debug for Builder {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.debug_struct("Builder").finish_non_exhaustive()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::PackURI;

    fn media_value(data: Vec<u8>) -> Result<crate::media_parts::List> {
        Ok(crate::media_parts::List {
            pictures: vec![crate::media_parts::Picture {
                shape_id: 4,
                name: "sample.mp4".into(),
                kind: crate::media_parts::Kind::Video,
                relationship_id: "rIdVideo".into(),
                resource: Some(crate::media_parts::Resource::new(
                    "/ppt/media/media1.mp4",
                    "video/mp4",
                    data,
                )),
                poster: None,
                transform: Some(crate::media_parts::Transform::emu(0, 0, 1, 1)?),
                office_extension: None,
            }],
        })
    }

    #[test]
    fn build_finalizes_authored_graph_for_immutable_edits() -> Result<()> {
        let mut builder = Builder::new()?;
        builder
            .presentation()?
            .add_slide()?
            .add_text_box("before", 0, 0, 914400, 914400);
        let package = builder.build()?;
        let mut original = Vec::new();
        package.write_to(&mut original)?;
        let mut edit = package.edit()?;
        assert!(edit.set_shape_text(0usize, 0usize, "after")?);
        let (changed, patch) = edit.commit()?.into_parts();
        assert!(!patch.is_noop());
        assert!(!package.same_content(&changed)?);
        let mut unchanged = Vec::new();
        package.write_to(&mut unchanged)?;
        assert_eq!(original, unchanged);
        let mut output = Vec::new();
        changed.write_to(&mut output)?;
        let reopened = Package::from_reader(std::io::Cursor::new(output))?;
        assert!(changed.same_content(&reopened)?);
        Ok(())
    }

    #[test]
    fn authored_publication_replaces_stale_content_identity() -> Result<()> {
        let mut builder = Builder::new()?;
        let _ = builder.package.content_revision.set([0xA5; 32]);
        builder
            .presentation()?
            .add_slide()?
            .add_text_box("published", 0, 0, 914_400, 914_400);

        let package = builder.build()?;

        assert_ne!(package.cached_content_revision(), Some([0xA5; 32]));
        assert!(package.part_bytes > 0);
        Ok(())
    }

    #[test]
    fn authored_notes_index_survives_changed_notes_replay() -> Result<()> {
        let mut builder = Builder::new()?;
        let presentation = builder.presentation()?;
        presentation.add_slide()?.set_notes("before");
        presentation.add_slide()?.set_notes("unchanged");
        let package = builder.build()?;
        let names = [
            "/ppt/notesMasters/notesMaster1.xml",
            "/ppt/theme/theme2.xml",
            "/ppt/notesSlides/notesSlide1.xml",
            "/ppt/notesSlides/notesSlide2.xml",
        ];
        let expected = names.into_iter().try_fold(0_u64, |total, name| {
            let part_name = PackURI::new(name).map_err(crate::Error::Uri)?;
            let bytes = u64::try_from(package.opc.get_part(&part_name)?.blob().len())
                .map_err(|_| crate::Error::Invalid("notes index byte count overflow".into()))?;
            total
                .checked_add(bytes)
                .ok_or_else(|| crate::Error::Invalid("notes index byte count overflow".into()))
        })?;
        assert_eq!(package.replay_index.notes_total_bytes(), expected);

        let mut edit = package.edit_notes()?;
        assert!(edit.set_text(0_usize, "after")?);
        let commit = edit.commit()?;
        let replayed = package.apply_notes(commit.patch())?;
        assert!(replayed.diagnostics().changed());
        Ok(())
    }

    #[test]
    fn authored_media_index_supports_resource_only_replay() -> Result<()> {
        let mut builder = Builder::new()?;
        builder.presentation()?.add_slide()?;
        let package = builder.build()?;
        let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(crate::Error::Uri)?;
        let initial = media_value(vec![0, 1, 2])?;
        let mut raw = package.edit_opc()?;
        raw.update(|opc| {
            crate::media_parts::store_with(
                opc,
                &slide_name,
                &initial,
                crate::media_parts::Conformance::Transitional,
                false,
            )
        })?;
        let with_media = raw.commit()?.into_package();
        let media_name = PackURI::new("/ppt/media/media1.mp4").map_err(crate::Error::Uri)?;
        assert!(
            with_media
                .replay_index
                .media_slides(&media_name)
                .is_some_and(|slides| slides.iter().any(|slide| slide == &slide_name))
        );

        let mut after = (*with_media.opc).clone();
        after
            .get_part_mut(&media_name)?
            .set_blob(vec![0, 1, 2, 3, 4]);
        let after = std::sync::Arc::new(after);
        let patch = crate::media_parts::Patch(crate::part_edit::PartPatch::capture(
            &with_media.opc,
            &after,
        )?);
        let replayed = with_media.apply_media(&patch)?;
        assert!(replayed.diagnostics().changed());
        Ok(())
    }
}
