//! Typed inert form/control reads and snapshot edits.

use super::super::super::model::MutableDocument;
use litchi_core::Result;

impl MutableDocument {
    /// Return all typed form/control custom properties in document order.
    pub fn form_properties(&self) -> Result<Vec<crate::form::Property>> {
        self.with_content_xml(crate::form::form_properties)
    }

    /// Insert a property into a form/control owner selected in document order.
    pub fn insert_form_property(
        &mut self,
        owner_index: usize,
        property: &crate::form::Property,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_form_property_xml(xml, owner_index, property)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a form property selected in document order.
    pub fn replace_form_property(
        &mut self,
        property_index: usize,
        replacement: &crate::form::Property,
    ) -> Result<crate::form::Property> {
        let old = self
            .form_properties()?
            .get(property_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "form property {property_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_form_property_xml(xml, property_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a form property and remove its container when it becomes empty.
    pub fn remove_form_property(&mut self, property_index: usize) -> Result<crate::form::Property> {
        let old = self
            .form_properties()?
            .get(property_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "form property {property_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::form::remove_form_property_xml(xml, property_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return text and textarea controls in document order.
    pub fn text_controls(&self) -> Result<Vec<crate::form::TextControl>> {
        self.with_content_xml(crate::form::text_controls)
    }

    /// Insert a text or textarea control into a form selected in document order.
    pub fn insert_text_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TextControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_text_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a text or textarea control selected in document order.
    pub fn replace_text_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TextControl,
    ) -> Result<crate::form::TextControl> {
        let old = self
            .text_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "text control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_text_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a text or textarea control selected in document order.
    pub fn remove_text_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::TextControl> {
        let old = self
            .text_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "text control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::form::remove_text_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return button and checkbox controls in document order.
    pub fn interactive_controls(&self) -> Result<Vec<crate::form::InteractiveControl>> {
        self.with_content_xml(crate::form::interactive_controls)
    }

    /// Insert a button or checkbox into a form selected in document order.
    pub fn insert_interactive_control(
        &mut self,
        form_index: usize,
        control: &crate::form::InteractiveControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_interactive_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a button or checkbox selected in document order.
    pub fn replace_interactive_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::InteractiveControl,
    ) -> Result<crate::form::InteractiveControl> {
        let old = self
            .interactive_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "interactive control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_interactive_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a button or checkbox selected in document order.
    pub fn remove_interactive_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::InteractiveControl> {
        let old = self
            .interactive_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "interactive control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_interactive_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return listbox and combobox controls in document order.
    pub fn selection_controls(&self) -> Result<Vec<crate::form::SelectionControl>> {
        self.with_content_xml(crate::form::selection_controls)
    }

    /// Insert a listbox or combobox into a form selected in document order.
    pub fn insert_selection_control(
        &mut self,
        form_index: usize,
        control: &crate::form::SelectionControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_selection_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a listbox or combobox selected in document order.
    pub fn replace_selection_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::SelectionControl,
    ) -> Result<crate::form::SelectionControl> {
        let old = self
            .selection_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "selection control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_selection_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a listbox or combobox selected in document order.
    pub fn remove_selection_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::SelectionControl> {
        let old = self
            .selection_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "selection control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_selection_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return radio, frame, and image-button controls in document order.
    pub fn visual_controls(&self) -> Result<Vec<crate::form::VisualControl>> {
        self.with_content_xml(crate::form::visual_controls)
    }

    /// Insert a radio, frame, or image-button into a form selected in document order.
    pub fn insert_visual_control(
        &mut self,
        form_index: usize,
        control: &crate::form::VisualControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_visual_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a radio, frame, or image-button selected in document order.
    pub fn replace_visual_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::VisualControl,
    ) -> Result<crate::form::VisualControl> {
        let old = self
            .visual_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "visual control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_visual_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a radio, frame, or image-button selected in document order.
    pub fn remove_visual_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::VisualControl> {
        let old = self
            .visual_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "visual control {control_index} is out of bounds"
                ))
            })?;
        let updated = self
            .with_content_xml(|xml| crate::form::remove_visual_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return fixed-text, hidden, and generic controls in document order.
    pub fn generic_form_controls(&self) -> Result<Vec<crate::form::GenericFormControl>> {
        self.with_content_xml(crate::form::generic_form_controls)
    }

    /// Insert a fixed-text, hidden, or generic control into a form selected in document order.
    pub fn insert_generic_form_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GenericFormControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_generic_form_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a fixed-text, hidden, or generic control selected in document order.
    pub fn replace_generic_form_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GenericFormControl,
    ) -> Result<crate::form::GenericFormControl> {
        let old = self
            .generic_form_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "generic form control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_generic_form_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a fixed-text, hidden, or generic control selected in document order.
    pub fn remove_generic_form_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::GenericFormControl> {
        let old = self
            .generic_form_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "generic form control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_generic_form_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return password and file controls in document order.
    pub fn password_file_controls(&self) -> Result<Vec<crate::form::PasswordFileControl>> {
        self.with_content_xml(crate::form::password_file_controls)
    }

    /// Insert a password or file control into a form selected in document order.
    pub fn insert_password_file_control(
        &mut self,
        form_index: usize,
        control: &crate::form::PasswordFileControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_password_file_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a password or file control selected in document order.
    pub fn replace_password_file_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::PasswordFileControl,
    ) -> Result<crate::form::PasswordFileControl> {
        let old = self
            .password_file_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "password/file control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_password_file_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a password or file control selected in document order.
    pub fn remove_password_file_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::PasswordFileControl> {
        let old = self
            .password_file_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "password/file control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_password_file_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return image-frame controls in document order without resolving image references.
    pub fn image_frame_controls(&self) -> Result<Vec<crate::form::ImageFrameControl>> {
        self.with_content_xml(crate::form::image_frame_controls)
    }

    /// Insert an image-frame control into a form selected in document order.
    pub fn insert_image_frame_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ImageFrameControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_image_frame_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace an image-frame control selected in document order.
    pub fn replace_image_frame_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ImageFrameControl,
    ) -> Result<crate::form::ImageFrameControl> {
        let old = self
            .image_frame_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "image-frame control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_image_frame_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove an image-frame control selected in document order.
    pub fn remove_image_frame_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::ImageFrameControl> {
        let old = self
            .image_frame_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "image-frame control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_image_frame_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return value-range controls in document order without resolving bindings.
    pub fn value_range_controls(&self) -> Result<Vec<crate::form::ValueRangeControl>> {
        self.with_content_xml(crate::form::value_range_controls)
    }

    /// Insert a value-range control into a form selected in document order.
    pub fn insert_value_range_control(
        &mut self,
        form_index: usize,
        control: &crate::form::ValueRangeControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_value_range_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a value-range control selected in document order.
    pub fn replace_value_range_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::ValueRangeControl,
    ) -> Result<crate::form::ValueRangeControl> {
        let old = self
            .value_range_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "value-range control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_value_range_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a value-range control selected in document order.
    pub fn remove_value_range_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::ValueRangeControl> {
        let old = self
            .value_range_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "value-range control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_value_range_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Return formatted-text, number, date, and time controls in document order.
    pub fn typed_value_controls(&self) -> Result<Vec<crate::form::TypedValueControl>> {
        self.with_content_xml(crate::form::typed_value_controls)
    }

    /// Insert a typed value control into a form selected in document order.
    pub fn insert_typed_value_control(
        &mut self,
        form_index: usize,
        control: &crate::form::TypedValueControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_typed_value_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }

    /// Replace a typed value control selected in document order.
    pub fn replace_typed_value_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::TypedValueControl,
    ) -> Result<crate::form::TypedValueControl> {
        let old = self
            .typed_value_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "typed value control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_typed_value_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    /// Remove a typed value control selected in document order.
    pub fn remove_typed_value_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::TypedValueControl> {
        let old = self
            .typed_value_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "typed value control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::remove_typed_value_control_xml(xml, control_index)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }

    pub fn grid_controls(&self) -> Result<Vec<crate::form::GridControl>> {
        self.with_content_xml(crate::form::grid_controls)
    }
    pub fn insert_grid_control(
        &mut self,
        form_index: usize,
        control: &crate::form::GridControl,
    ) -> Result<()> {
        let updated = self.with_content_xml(|xml| {
            crate::form::insert_grid_control_xml(xml, form_index, control)
        })?;
        self.content_xml = Some(updated);
        Ok(())
    }
    pub fn replace_grid_control(
        &mut self,
        control_index: usize,
        replacement: &crate::form::GridControl,
    ) -> Result<crate::form::GridControl> {
        let old = self
            .grid_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "grid control {control_index} is out of bounds"
                ))
            })?;
        let updated = self.with_content_xml(|xml| {
            crate::form::replace_grid_control_xml(xml, control_index, replacement)
        })?;
        self.content_xml = Some(updated);
        Ok(old)
    }
    pub fn remove_grid_control(
        &mut self,
        control_index: usize,
    ) -> Result<crate::form::GridControl> {
        let old = self
            .grid_controls()?
            .get(control_index)
            .cloned()
            .ok_or_else(|| {
                litchi_core::Error::InvalidFormat(format!(
                    "grid control {control_index} is out of bounds"
                ))
            })?;
        let updated =
            self.with_content_xml(|xml| crate::form::remove_grid_control_xml(xml, control_index))?;
        self.content_xml = Some(updated);
        Ok(old)
    }
}
