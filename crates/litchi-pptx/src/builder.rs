//! Typed presentation authoring, finalized as a canonical package.
use crate::writer::MutablePresentation;
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
}
impl Builder {
    /// Start the native minimal empty presentation.
    ///
    /// # Errors
    /// Returns native construction failures.
    pub fn new() -> Result<Self> {
        Ok(Self {
            package: Package::new_authoring()?,
        })
    }

    /// Configure the native typed presentation before finalization.
    ///
    /// # Errors
    /// Returns native authoring admission failures.
    pub fn presentation(&mut self) -> Result<&mut MutablePresentation> {
        self.package.authoring_presentation()
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
        self.package.set_font_embedding(mode)?;
        Ok(self)
    }

    /// Finalize the native graph without serializing and reopening a ZIP.
    ///
    /// # Errors
    /// Returns native graph construction and validation failures.
    pub fn build(self) -> Result<Package> {
        self.package.finish_creation()
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
    #[test]
    fn build_finalizes_authored_graph_for_immutable_edits() -> Result<()> {
        let mut builder = Builder::new()?;
        builder
            .presentation()?
            .add_slide()?
            .add_text_box("before", 0, 0, 914400, 914400);
        let package = builder.build()?;
        assert!(!package.is_modified());
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
        let reopened = Package::from_vec(output)?;
        assert!(changed.same_content(&reopened)?);
        Ok(())
    }
}
