//! Message-type dispatch for generated protobuf messages.

use super::wrappers::*;
use super::{kn, tn, tp, tsch, tsd, tsp, tst, tswp};
use crate::{Error, Result};
use phf::phf_map;
use prost::Message;

macro_rules! decoder {
    ($name:ident, $message:path, $wrapper:ident) => {
        fn $name(data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
            Ok(Box::new($wrapper(<$message>::decode(data)?)))
        }
    };
}

decoder!(decode_archive_info, tsp::ArchiveInfo, ArchiveInfoWrapper);
decoder!(decode_message_info, tsp::MessageInfo, MessageInfoWrapper);
decoder!(
    decode_pages_document,
    tp::DocumentArchive,
    PagesDocumentWrapper
);
decoder!(decode_numbers_sheet, tn::SheetArchive, NumbersSheetWrapper);
decoder!(decode_keynote_slide, kn::SlideArchive, KeynoteSlideWrapper);
decoder!(
    decode_storage_archive,
    tswp::StorageArchive,
    StorageArchiveWrapper
);
decoder!(
    decode_table_model,
    tst::TableModelArchive,
    TableModelWrapper
);
decoder!(
    decode_table_data_list,
    tst::TableDataList,
    TableDataListWrapper
);
decoder!(decode_shape_archive, tsd::ShapeArchive, ShapeArchiveWrapper);
decoder!(
    decode_drawable_archive,
    tsd::DrawableArchive,
    DrawableArchiveWrapper
);
decoder!(
    decode_chart_archive,
    tsch::ChartArchive,
    ChartArchiveWrapper
);

type Decoder = fn(&[u8]) -> Result<Box<dyn DecodedMessage>>;

static DECODERS: phf::Map<u32, Decoder> = phf_map! {
    1u32 => decode_archive_info,
    2u32 => decode_message_info,
    3u32 => decode_numbers_sheet,
    5u32 => decode_keynote_slide,
    6u32 => decode_keynote_slide,
    8u32 => decode_keynote_slide,
    2001u32 => decode_storage_archive,
    2002u32 => decode_storage_archive,
    2003u32 => decode_storage_archive,
    2004u32 => decode_storage_archive,
    2005u32 => decode_storage_archive,
    2006u32 => decode_storage_archive,
    2007u32 => decode_storage_archive,
    2008u32 => decode_storage_archive,
    2009u32 => decode_storage_archive,
    2010u32 => decode_storage_archive,
    2011u32 => decode_storage_archive,
    2012u32 => decode_storage_archive,
    2013u32 => decode_storage_archive,
    2014u32 => decode_storage_archive,
    2022u32 => decode_storage_archive,
    3002u32 => decode_drawable_archive,
    3003u32 => decode_drawable_archive,
    3004u32 => decode_shape_archive,
    3005u32 => decode_shape_archive,
    3006u32 => decode_shape_archive,
    3007u32 => decode_shape_archive,
    3008u32 => decode_shape_archive,
    3009u32 => decode_shape_archive,
    5000u32 => decode_chart_archive,
    5004u32 => decode_chart_archive,
    5021u32 => decode_chart_archive,
    6000u32 => decode_table_model,
    6001u32 => decode_table_model,
    6005u32 => decode_table_data_list,
    6201u32 => decode_table_data_list,
    10000u32 => decode_pages_document,
    10001u32 => decode_pages_document,
    10011u32 => decode_pages_document,
};

/// Decode a message using its numeric iWork message type.
pub fn decode(message_type: u32, data: &[u8]) -> Result<Box<dyn DecodedMessage>> {
    DECODERS
        .get(&message_type)
        .ok_or(Error::UnsupportedMessageType(message_type))?(data)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn expected_decoders_are_registered() {
        for message_type in [1, 2, 3, 5, 6, 2001, 2002, 2003, 2022, 6001, 10000] {
            assert!(DECODERS.contains_key(&message_type));
        }
    }

    #[test]
    fn unsupported_message_type_is_reported() {
        assert!(matches!(
            decode(999, &[]),
            Err(Error::UnsupportedMessageType(999))
        ));
    }

    #[test]
    fn registered_decoders_are_callable() {
        for message_type in [1, 2, 3, 5, 2001, 6001, 10000] {
            assert!(decode(message_type, &[0; 10]).is_err());
        }
    }
}
