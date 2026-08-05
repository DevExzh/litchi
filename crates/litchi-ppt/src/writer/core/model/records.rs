//! Typed record and snapshot planning for the PPT writer model.

use super::{SerializedHeaderFooters, WriteError, Writer};
use crate::encryption::{WriterEncryptionMaterial, prepare_writer_encryption};
use crate::header_footer::{
    HeaderFooter, HeaderFooterParent, HeaderFooterParentOrdinal, HeaderFooterScope,
};
use crate::modify_password::validate_value as validate_modify_password;
use crate::writer::chart::ChartPlan;
use crate::writer::persist::PersistPtrBuilder;
use crate::writer::records::{RecordBuilder, record_type};
use crate::writer::spec::Tag10;

impl Writer {
    pub(in crate::writer::core) fn build_modify_password_programmable_tag(
        &self,
    ) -> Result<Option<Vec<u8>>, WriteError> {
        let Some(value) = &self.modify_password else {
            return Ok(None);
        };
        validate_modify_password(value.password.as_str())
            .map_err(|error| WriteError::InvalidData(error.to_string()))?;

        let mut atom = RecordBuilder::new(0x00, 3, record_type::CSTRING);
        let atom_data = value
            .password
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>();
        atom.write_data(&atom_data);
        let mut blob = RecordBuilder::new(0x00, 0, record_type::BINARY_TAG_DATA);
        blob.write_child(&atom.build()?);
        let mut binary_tag = RecordBuilder::new(0x0f, 0, record_type::PROG_BINARY_TAG);
        let mut name = RecordBuilder::new(0x00, 0, record_type::CSTRING);
        name.write_data(&Tag10::to_bytes());
        binary_tag.write_child(&name.build()?);
        binary_tag.write_child(&blob.build()?);
        let mut tags = RecordBuilder::new(0x0f, 0, record_type::PROG_TAGS);
        tags.write_child(&binary_tag.build()?);
        Ok(Some(tags.build()?))
    }

    pub(in crate::writer::core) fn prepare_encryption(
        &self,
    ) -> Result<Option<WriterEncryptionMaterial>, WriteError> {
        self.encryption
            .as_ref()
            .map(|value| prepare_writer_encryption(value.profile, value.password.as_str()))
            .transpose()
            .map_err(WriteError::InvalidData)
    }
}

impl Writer {
    pub(in crate::writer::core) fn serialize_header_footers(
        &self,
    ) -> Result<SerializedHeaderFooters, WriteError> {
        let serialize = |value: Option<&HeaderFooter>, scope| {
            value
                .map(|value| {
                    let mut value = value.clone();
                    value.scope = scope;
                    value
                        .to_record_bytes()
                        .map_err(|error| WriteError::InvalidData(error.to_string()))
                })
                .transpose()
        };
        let presentation_slides = serialize(
            self.presentation_header_footer.as_ref(),
            HeaderFooterScope::PresentationSlides,
        )?;
        let notes_and_handouts = serialize(
            self.notes_and_handouts_header_footer.as_ref(),
            HeaderFooterScope::NotesAndHandouts,
        )?;
        let main_master = serialize(
            self.main_master_header_footer.as_ref(),
            HeaderFooterScope::Local {
                parent: HeaderFooterParent::MainMaster,
                parent_ordinal: HeaderFooterParentOrdinal::new(0),
            },
        )?;
        let slides = self
            .slides
            .iter()
            .enumerate()
            .map(|(index, slide)| {
                serialize(
                    slide.header_footer.as_ref(),
                    HeaderFooterScope::Local {
                        parent: HeaderFooterParent::Slide,
                        parent_ordinal: HeaderFooterParentOrdinal::new(index),
                    },
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        Ok(SerializedHeaderFooters {
            presentation_slides,
            notes_and_handouts,
            main_master,
            slides,
        })
    }
}

impl Writer {
    /// Assign external-object and persist identifiers to every chart.
    ///
    /// Chart object identifiers continue above the hyperlink identifier seed
    /// because both share the `ExObjId` namespace ([MS-PPT] 2.10.1).
    pub(in crate::writer::core) fn plan_charts(
        &self,
        persist_builder: &mut PersistPtrBuilder,
    ) -> Result<Vec<ChartPlan>, WriteError> {
        let total: usize = self.slides.iter().map(|slide| slide.charts.len()).sum();
        if total > crate::writer::chart::MAX_CHART_OBJECTS {
            return Err(WriteError::InvalidData(format!(
                "presentation exceeds {} chart objects",
                crate::writer::chart::MAX_CHART_OBJECTS
            )));
        }
        let mut next_id = self.hyperlinks.id_seed();
        let mut plans = Vec::with_capacity(total);
        for (slide_index, slide) in self.slides.iter().enumerate() {
            for chart_index in 0..slide.charts.len() {
                next_id = next_id.checked_add(1).ok_or_else(|| {
                    WriteError::InvalidData("external-object ID space exhausted".to_string())
                })?;
                plans.push(ChartPlan {
                    slide: slide_index,
                    chart: chart_index,
                    ex_obj_id: next_id,
                    persist_id: persist_builder.allocate_id(),
                });
            }
        }
        Ok(plans)
    }
}
