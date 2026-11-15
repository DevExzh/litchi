//! Slide master and slide layout names and placeholder text.
//!
//! Masters are selected by their zero-based position in
//! `p:sldMasterIdLst`, layouts by their position in their master's
//! `p:sldLayoutIdLst`, as the read catalog numbers them. Each verb rewrites
//! only the attribute or the text runs it names, keeping every other byte
//! of the part.

use litchi_core::Position;
use litchi_opc::PackURI;
use litchi_opc::constants::content_type as ct;

use super::{Edit, invalid};
use crate::Result;

impl Edit {
    /// Rename a slide master (`p:cSld/@name`). An equal name is an exact
    /// no-op. Returns whether the master changed.
    ///
    /// ```
    /// let source = litchi_pptx::Package::new()?;
    /// let mut edit = source.edit()?;
    /// assert!(edit.set_master_name(litchi_core::Position::new(0), "Brand")?);
    /// let package = edit.commit()?.into_package();
    /// let presentation = package.presentation()?;
    /// let master = &presentation.slide_masters()?[0];
    /// assert_eq!(master.part().authored_name()?.as_deref(), Some("Brand"));
    /// # Ok::<(), litchi_pptx::Error>(())
    /// ```
    ///
    /// # Errors
    ///
    /// Returns an error for a missing master, a name XML cannot carry, or a
    /// malformed master part.
    pub fn set_master_name(&mut self, master: Position, name: &str) -> Result<bool> {
        self.ensure_ready()?;
        let part = self.master_part(master)?;
        self.set_part_name(&part, ct::PML_SLIDE_MASTER, b"sldMaster", name)
    }

    /// Rename a slide layout (`p:cSld/@name`) of a master. An equal name is
    /// an exact no-op. Returns whether the layout changed.
    ///
    /// Slides refer to their layout by relationship, not by name, so a
    /// rename keeps every slide bound; [`Edit::insert_slide`] looks layouts
    /// up by their new name afterwards.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing master or layout, a name XML cannot
    /// carry, or a malformed layout part.
    pub fn set_layout_name(
        &mut self,
        master: Position,
        layout: Position,
        name: &str,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let part = self.layout_part(master, layout)?;
        self.set_part_name(&part, ct::PML_SLIDE_LAYOUT, b"sldLayout", name)
    }

    /// Replace the text of one shape (a placeholder, typically) of a slide
    /// master, keeping its first run's formatting, as
    /// [`Edit::set_shape_text`] does on a slide. Returns whether the master
    /// changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing master or shape, a shape without a
    /// text body, or text XML cannot carry, and a refusal for a shape that
    /// holds an Office Math equation.
    pub fn set_master_shape_text<'k>(
        &mut self,
        master: Position,
        shape: impl Into<crate::shape::Key<'k>>,
        text: &str,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let part = self.master_part(master)?;
        self.set_part_shape_text(
            &part,
            ct::PML_SLIDE_MASTER,
            shape.into(),
            text,
            "set_master_shape_text",
        )
    }

    /// Replace the text of one shape of a slide layout, as
    /// [`Edit::set_master_shape_text`]. Returns whether the layout changed.
    ///
    /// # Errors
    ///
    /// Returns an error for a missing master, layout or shape, a shape
    /// without a text body, or text XML cannot carry, and a refusal for a
    /// shape that holds an Office Math equation.
    pub fn set_layout_shape_text<'k>(
        &mut self,
        master: Position,
        layout: Position,
        shape: impl Into<crate::shape::Key<'k>>,
        text: &str,
    ) -> Result<bool> {
        self.ensure_ready()?;
        let part = self.layout_part(master, layout)?;
        self.set_part_shape_text(
            &part,
            ct::PML_SLIDE_LAYOUT,
            shape.into(),
            text,
            "set_layout_shape_text",
        )
    }

    fn master_part(&self, master: Position) -> Result<PackURI> {
        let presentation = crate::presentation::Presentation::new_with_semantic_limits(
            crate::parts::PresentationPart::from_package(self.working_ref()?)?,
            self.working_ref()?,
            self.semantic_limits_ref(),
        );
        let masters = presentation.slide_masters()?;
        let selected = masters
            .get(master.get())
            .ok_or_else(|| invalid("the slide master position is past the last master"))?;
        Ok(selected.part().part().partname().clone())
    }

    fn layout_part(&self, master: Position, layout: Position) -> Result<PackURI> {
        let presentation = crate::presentation::Presentation::new_with_semantic_limits(
            crate::parts::PresentationPart::from_package(self.working_ref()?)?,
            self.working_ref()?,
            self.semantic_limits_ref(),
        );
        let masters = presentation.slide_masters()?;
        let selected = masters
            .get(master.get())
            .ok_or_else(|| invalid("the slide master position is past the last master"))?;
        let layouts = selected.layouts()?;
        let selected = layouts
            .get(layout.get())
            .ok_or_else(|| invalid("the slide layout position is past the master's last layout"))?;
        Ok(selected.part().part().partname().clone())
    }

    fn set_part_name(
        &mut self,
        part: &PackURI,
        content_type: &str,
        root: &[u8],
        name: &str,
    ) -> Result<bool> {
        let owner = self.working.get_part(part)?;
        crate::parts::validate_content_type(owner, content_type)?;
        let Some(xml) = super::super::stage_common_slide_name(
            owner.blob(),
            root,
            name,
            self.semantic_limits().shape().output_bytes(),
        )?
        else {
            return Ok(false);
        };
        self.working.get_part_mut(part)?.set_blob(xml);
        Ok(true)
    }

    fn set_part_shape_text(
        &mut self,
        part: &PackURI,
        content_type: &str,
        key: crate::shape::Key<'_>,
        text: &str,
        operation: &'static str,
    ) -> Result<bool> {
        let owner = self.working.get_part(part)?;
        crate::parts::validate_content_type(owner, content_type)?;
        let shape_limits = self.semantic_limits().shape();
        let scene = crate::shape::Scene::read(owner.blob(), shape_limits)?;
        let selected = scene.shape(key)?;
        if selected.common().text() == Some(text) {
            return Ok(false);
        }
        if selected.common().text().is_none() {
            return Err(invalid("the selected shape has no text body"));
        }
        let span = crate::tag::shape::selected_raw_span_for_shape(owner.blob(), &scene, selected)?;
        super::refuse_equation_shape(
            owner.blob(),
            &scene,
            &span,
            operation,
            super::EQUATION_SHAPE_REASON,
        )?;
        let xml = super::super::xml::rewrite_shape_text_with_limits(
            owner.blob(),
            span,
            text,
            self.source_ref()?.limits.max_text_bytes(),
            shape_limits,
            super::super::xml::TextRefusals::SHAPE,
        )?;
        let staged = crate::shape::Scene::read(&xml, shape_limits)?;
        if staged.shape(key)?.common().text() != Some(text) || staged.len() != scene.len() {
            return Err(invalid("the shape text did not round-trip semantically"));
        }
        self.working.get_part_mut(part)?.set_blob(xml);
        Ok(true)
    }
}

#[cfg(test)]
mod tests {
    use litchi_core::Position;

    #[test]
    fn master_and_layout_names_and_text_round_trip() {
        let source = crate::Package::new().unwrap();
        let mut edit = source.edit().unwrap();
        let first = Position::new(0);
        assert!(edit.set_master_name(first, "Brand").unwrap());
        assert!(!edit.set_master_name(first, "Brand").unwrap());
        assert!(edit.set_layout_name(first, first, "Opening").unwrap());
        assert!(edit.set_master_name(Position::new(9), "x").is_err());
        let package = edit.commit().unwrap().into_package();
        let presentation = package.presentation().unwrap();
        let masters = presentation.slide_masters().unwrap();
        assert_eq!(
            masters[0].part().authored_name().unwrap().as_deref(),
            Some("Brand")
        );
        assert_eq!(
            masters[0].layouts().unwrap()[0]
                .part()
                .authored_name()
                .unwrap()
                .as_deref(),
            Some("Opening")
        );
    }

    #[test]
    fn master_equation_shape_text_is_refused_and_ordinary_shapes_still_edit() {
        let source = crate::Package::new().unwrap();
        let mut opc = litchi_opc::OpcPackage::from_bytes(
            &source.to_bytes().unwrap(),
            litchi_opc::SourceLimits::default(),
        )
        .unwrap();
        let uri = litchi_opc::PackURI::new("/ppt/slideMasters/slideMaster1.xml").unwrap();
        let part = opc.get_part_mut(&uri).unwrap();
        let xml = std::str::from_utf8(part.blob()).unwrap().replacen(
            "</p:spTree>",
            concat!(
                r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main">"#,
                r#"<mc:Choice Requires="a14"><p:sp><p:nvSpPr><p:cNvPr id="90" name="Equation"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>"#,
                r#"<p:txBody><a:bodyPr/><a:p><a14:m><m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath></a14:m></a:p></p:txBody></p:sp></mc:Choice>"#,
                r#"<mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="90" name="Equation"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/>"#,
                r#"<p:txBody><a:bodyPr/><a:p><a:r><a:t> </a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent></p:spTree>"#,
            ),
            1,
        );
        part.set_blob(xml.into_bytes());
        let package = crate::Package::from_opc_package(opc).unwrap();
        let master_xml = |package: &crate::Package| {
            package
                .opc()
                .unwrap()
                .get_part(&uri)
                .unwrap()
                .blob()
                .to_vec()
        };
        let original = master_xml(&package);
        let mut edit = package.edit().unwrap();
        let first = Position::new(0);
        let error = edit
            .set_master_shape_text(first, "Equation", "EDITED")
            .unwrap_err();
        let refusal = error.refusal().expect("a typed refusal");
        assert_eq!(refusal.operation(), "set_master_shape_text");
        assert_eq!(
            refusal.reason(),
            "the shape holds an Office Math equation, which text edits cannot change"
        );
        let unchanged = edit.commit().unwrap().into_package();
        assert_eq!(master_xml(&unchanged), original);
        let mut edit = unchanged.edit().unwrap();
        assert!(
            edit.set_master_shape_text(first, 0_usize, "EDITED")
                .unwrap()
        );
        let edited = master_xml(&edit.commit().unwrap().into_package());
        assert!(std::str::from_utf8(&edited).unwrap().contains("EDITED"));
    }
}
