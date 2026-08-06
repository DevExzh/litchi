//! Slide-level timing extension scanning.

use super::parser::parse_extended_time_node;
use crate::animation::linked_slide::{LinkedShape, LinkedSlide};
use crate::animation::slide_metadata::SlideTime;
use crate::animation::types::{Flags, SlideAnimationExtension};
use crate::consts::RecordType;
use crate::package::{Error, Result};
use crate::records::Record;

/// Discover PowerPoint 2002 timing and build records in `BinaryTagData`.
pub fn parse_slide_animation_extension(data: &[u8]) -> Result<SlideAnimationExtension> {
    let mut extension = SlideAnimationExtension::default();
    let mut offset = 0usize;
    let mut linked_shape_array_closed = false;
    while offset < data.len() {
        if data.len() - offset < 8 {
            return Err(Error::Corrupted(
                "slide binary tag ends with a partial record header".to_string(),
            ));
        }
        let (record, consumed) = Record::parse(data, offset)?;
        if record.data_length as usize != record.data.len() || consumed < 8 {
            return Err(Error::Corrupted(format!(
                "slide binary tag contains a truncated {:?} record",
                record.record_type
            )));
        }
        if record.record_type != RecordType::LinkedShape10Atom
            && !extension.linked_shapes.is_empty()
        {
            linked_shape_array_closed = true;
        }
        match record.record_type {
            RecordType::LinkedSlide10Atom => {
                if extension.linked_slide.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple LinkedSlide10Atom records".to_string(),
                    ));
                }
                if !extension.linked_shapes.is_empty() {
                    return Err(Error::InvalidFormat(
                        "LinkedSlide10Atom must precede its LinkedShape10Atom array".to_string(),
                    ));
                }
                extension.linked_slide = Some(LinkedSlide::parse_record(&record)?);
            },
            RecordType::LinkedShape10Atom => {
                let linked_slide = extension.linked_slide.ok_or_else(|| {
                    Error::InvalidFormat(
                        "LinkedShape10Atom requires a preceding LinkedSlide10Atom".to_string(),
                    )
                })?;
                if linked_shape_array_closed {
                    return Err(Error::InvalidFormat(
                        "LinkedShape10Atom array must be contiguous".to_string(),
                    ));
                }
                let declared_count =
                    usize::try_from(linked_slide.linked_shape_count()).map_err(|_| {
                        Error::InvalidFormat(
                            "LinkedSlide10Atom shape count does not fit this platform".to_string(),
                        )
                    })?;
                if extension.linked_shapes.len() >= declared_count {
                    return Err(Error::InvalidFormat(
                        "LinkedShape10Atom array exceeds its declared count".to_string(),
                    ));
                }
                extension
                    .linked_shapes
                    .push(LinkedShape::parse_record(&record)?);
            },
            RecordType::ExtTimeNode => {
                if extension.time_node.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple root ExtTimeNode records".to_string(),
                    ));
                }
                extension.time_node = Some(parse_extended_time_node(&record)?);
            },
            RecordType::BuildList => {
                if extension.build_list.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple BuildList records".to_string(),
                    ));
                }
                extension.build_list = Some(super::super::build::parse_build_list(&record)?);
            },
            RecordType::SlideFlags10Atom => {
                if extension.slide_flags.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple SlideFlags10Atom records".to_string(),
                    ));
                }
                extension.slide_flags = Some(Flags::parse_record(&record)?);
            },
            RecordType::SlideTime10Atom => {
                if extension.creation_time_filetime.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple SlideTime10Atom records".to_string(),
                    ));
                }
                extension.creation_time_filetime =
                    Some(SlideTime::parse_record(&record)?.file_time());
            },
            RecordType::HashCode10Atom => {
                if extension.animation_hash.is_some() {
                    return Err(Error::InvalidFormat(
                        "___PPT10 contains multiple HashCode10Atom records".to_string(),
                    ));
                }
                extension.animation_hash =
                    Some(crate::animation::hash::Hash10::parse_record(&record)?.hash());
            },
            _ => {},
        }
        offset = offset
            .checked_add(consumed)
            .ok_or_else(|| Error::Corrupted("slide binary tag offset overflow".to_string()))?;
    }
    if let Some(linked_slide) = extension.linked_slide {
        let declared_count = usize::try_from(linked_slide.linked_shape_count()).map_err(|_| {
            Error::InvalidFormat(
                "LinkedSlide10Atom shape count does not fit this platform".to_string(),
            )
        })?;
        if extension.linked_shapes.len() != declared_count {
            return Err(Error::InvalidFormat(format!(
                "LinkedSlide10Atom declares {declared_count} linked shapes but {} were present",
                extension.linked_shapes.len()
            )));
        }
    }
    Ok(extension)
}
