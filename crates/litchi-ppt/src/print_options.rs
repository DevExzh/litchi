//! Typed PowerPoint document print preferences.

use super::package::{Error, Result};
use super::records::Record;
use crate::consts::RecordType;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u8)]
pub enum PrintTarget {
    Slides = 0,
    BuildSlides = 1,
    Handouts2 = 2,
    Handouts3 = 3,
    Handouts6 = 4,
    Notes = 5,
    Outline = 6,
    Handouts4 = 7,
    Handouts9 = 8,
    Handouts1 = 9,
}

impl PrintTarget {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::Slides),
            1 => Ok(Self::BuildSlides),
            2 => Ok(Self::Handouts2),
            3 => Ok(Self::Handouts3),
            4 => Ok(Self::Handouts6),
            5 => Ok(Self::Notes),
            6 => Ok(Self::Outline),
            7 => Ok(Self::Handouts4),
            8 => Ok(Self::Handouts9),
            9 => Ok(Self::Handouts1),
            _ => corrupted(format!("invalid PrintWhatEnum value {value:#04x}")),
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[repr(u8)]
pub enum PrintColorMode {
    BlackAndWhite = 0,
    Grayscale = 1,
    Color = 2,
}

impl PrintColorMode {
    fn parse(value: u8) -> Result<Self> {
        match value {
            0 => Ok(Self::BlackAndWhite),
            1 => Ok(Self::Grayscale),
            2 => Ok(Self::Color),
            _ => corrupted(format!("invalid ColorModeEnum value {value:#04x}")),
        }
    }
}

/// A validated document-level `PrintOptionsAtom`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct PrintOptions {
    pub target: PrintTarget,
    pub color_mode: PrintColorMode,
    pub print_hidden_slides: bool,
    pub scale_to_fit_paper: bool,
    pub frame_slides: bool,
}

impl PrintOptions {
    pub fn parse(document: &Record) -> Result<Option<Self>> {
        let records = document
            .children
            .iter()
            .filter(|record| record.record_type_raw == RecordType::PrintOptionsAtom.as_u16())
            .collect::<Vec<_>>();
        if records.len() > 1 {
            return corrupted("DocumentContainer contains duplicate PrintOptionsAtom records");
        }
        let Some(record) = records.first() else {
            return Ok(None);
        };
        if record.version != 0
            || record.instance != 0
            || record.data.len() != 5
            || record.data_length != 5
        {
            return corrupted("PrintOptionsAtom has an invalid header or size");
        }
        Ok(Some(Self {
            target: PrintTarget::parse(record.data[0])?,
            color_mode: PrintColorMode::parse(record.data[1])?,
            print_hidden_slides: parse_bool1(record.data[2], "fPrintHidden")?,
            scale_to_fit_paper: parse_bool1(record.data[3], "fScaleToFitPaper")?,
            frame_slides: parse_bool1(record.data[4], "fFrameSlides")?,
        }))
    }

    pub fn to_record(&self) -> Result<Record> {
        let bytes = self.to_record_bytes();
        let (record, end) = Record::parse(&bytes, 0)?;
        if end != bytes.len() {
            return corrupted("canonical PrintOptionsAtom did not consume its bytes");
        }
        Ok(record)
    }

    pub fn to_record_bytes(&self) -> [u8; 13] {
        let mut bytes = [0; 13];
        bytes[2..4].copy_from_slice(&RecordType::PrintOptionsAtom.as_u16().to_le_bytes());
        bytes[4..8].copy_from_slice(&5u32.to_le_bytes());
        bytes[8] = self.target as u8;
        bytes[9] = self.color_mode as u8;
        bytes[10] = self.print_hidden_slides.into();
        bytes[11] = self.scale_to_fit_paper.into();
        bytes[12] = self.frame_slides.into();
        bytes
    }
}

fn parse_bool1(value: u8, field: &str) -> Result<bool> {
    match value {
        0 => Ok(false),
        1 => Ok(true),
        _ => corrupted(format!("PrintOptionsAtom {field} is not bool1")),
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn root(children: Vec<Record>) -> Record {
        Record {
            version: 0x0f,
            instance: 0,
            record_type: RecordType::Document,
            record_type_raw: RecordType::Document.as_u16(),
            data_length: 0,
            data: Vec::new(),
            children,
        }
    }

    #[test]
    fn all_print_targets_and_color_modes_roundtrip() {
        for target in [
            PrintTarget::Slides,
            PrintTarget::BuildSlides,
            PrintTarget::Handouts2,
            PrintTarget::Handouts3,
            PrintTarget::Handouts6,
            PrintTarget::Notes,
            PrintTarget::Outline,
            PrintTarget::Handouts4,
            PrintTarget::Handouts9,
            PrintTarget::Handouts1,
        ] {
            for color_mode in [
                PrintColorMode::BlackAndWhite,
                PrintColorMode::Grayscale,
                PrintColorMode::Color,
            ] {
                let expected = PrintOptions {
                    target,
                    color_mode,
                    print_hidden_slides: true,
                    scale_to_fit_paper: false,
                    frame_slides: true,
                };
                assert_eq!(
                    PrintOptions::parse(&root(vec![expected.to_record().unwrap()])).unwrap(),
                    Some(expected)
                );
            }
        }
    }

    #[test]
    fn rejects_duplicate_invalid_header_enums_and_bool1() {
        let value = PrintOptions {
            target: PrintTarget::Slides,
            color_mode: PrintColorMode::Color,
            print_hidden_slides: false,
            scale_to_fit_paper: false,
            frame_slides: false,
        };
        let record = value.to_record().unwrap();
        assert!(PrintOptions::parse(&root(vec![record.clone(), record])).is_err());
        for (offset, invalid) in [(0, 1), (8, 10), (9, 3), (10, 2), (11, 0xff), (12, 7)] {
            let mut bytes = value.to_record_bytes();
            bytes[offset] = invalid;
            let record = Record::parse(&bytes, 0).unwrap().0;
            assert!(PrintOptions::parse(&root(vec![record])).is_err());
        }
    }
}
