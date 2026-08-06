//! Semantic collection of workbook-global BIFF records.

use super::super::WorkbookGlobalsSink;
use crate::defined_names::{DefinedNameSlot, LBL_RECORD_TYPE, NAME_CMT_RECORD_TYPE};
use crate::error::{Error, Result};
use crate::number_format::{DateSystem, Formatting};
use crate::protection;
use crate::records::{BofRecord, BoundSheetRecord, Encoding, SharedStringTable};
use crate::workbook::model::Workbook;
use litchi_biff::Records;
use std::io::{Read, Seek};
use std::sync::Arc;

impl<R: Read + Seek> Workbook<R> {
    /// Parse workbook globals (SST, bound sheets, etc.)
    pub(crate) fn parse_workbook_globals(
        &mut self,
        records_iter: &mut Records<'_>,
        encoding: &mut Encoding,
        sink: WorkbookGlobalsSink<'_>,
    ) -> Result<()> {
        let WorkbookGlobalsSink {
            bound_sheets,
            strings,
            string_properties,
            defined_name_slots,
            tolerance,
        } = sink;
        // Collect all records first for easier processing
        let mut records = Vec::new();
        for record_result in records_iter.by_ref() {
            let record = record_result?;
            let is_globals_eof = record.kind().get() == 0x000A;
            records.push(record);
            if is_globals_eof {
                break;
            }
        }

        self.formatting = Arc::new(Formatting::parse_globals(&records, tolerance)?);
        self.is_1904_date_system = self.formatting.date_system() == DateSystem::Excel1904;

        let mut palette_seen = false;
        let mut protection_collector = protection::WorkbookProtectionCollector::new();
        let mut calculation_collector = crate::calculation::WorkbookCalculationCollector::new();
        let mut vba_collector = crate::vba::WorkbookVbaCollector::new();
        let mut environment_collector = crate::environment::EnvironmentCollector::new();
        let mut write_access_collector = crate::access::WriteAccessCollector::new();
        let mut table_styles_collector = crate::table_styles::TableStylesCollector::new();
        let mut shared_string_index_collector =
            crate::shared_string_index::SharedStringIndexCollector::new();
        let mut workbook_view_collector = crate::workbook_view::WorkbookViewCollector::new();
        let mut function_group_collector = crate::function_group::FunctionGroupCollector::new();
        let mut external_link_collector = crate::external_link::ExternalLinkCollector::new();
        let mut name_optional_target: Option<(usize, u8)> = None;
        let mut i = 0;
        while i < records.len() {
            let record = &records[i];
            if !matches!(
                record.kind().get(),
                LBL_RECORD_TYPE
                    | NAME_CMT_RECORD_TYPE
                    | crate::defined_names::NAME_FN_GRP12_RECORD_TYPE
                    | crate::defined_names::NAME_PUBLISH_RECORD_TYPE
                    | 0x003c
            ) {
                name_optional_target = None;
            }
            if record.kind().get() == 0x003c && name_optional_target.is_some() {
                return Err(Error::InvalidRecord {
                    record_type: 0x003c,
                    message: "CONTINUE is not permitted after a post-Lbl optional record"
                        .to_string(),
                });
            }
            protection_collector.feed_record(record.kind().get(), record.payload())?;
            calculation_collector.feed_record(record.kind().get(), record.payload())?;
            vba_collector.feed_record(record.kind().get(), record.payload())?;
            environment_collector.feed_record(record.kind().get(), record.payload())?;
            write_access_collector.feed_record(record.kind().get(), record.payload());
            table_styles_collector.feed_record(record.kind().get(), record.payload())?;
            shared_string_index_collector.feed_record(record.kind().get(), record.payload());
            workbook_view_collector.feed_record(record.kind().get(), record.payload())?;
            function_group_collector.feed_record(record.kind().get(), record.payload())?;
            external_link_collector.feed_record(record.kind().get(), record.payload())?;

            match record.kind().get() {
                crate::font::FONT_RECORD_TYPE => {
                    let index = crate::font::logical_font_index(self.fonts.len())?;
                    self.fonts
                        .push(crate::font::Font::parse_record(
                            index,
                            record.payload(),
                            tolerance,
                        )?);
                },
                0x0092 => {
                    self.palette = crate::palette::Palette::parse_unique_record(
                        record.payload(),
                        &mut palette_seen,
                    )?;
                },
                crate::theme::THEME_RECORD_TYPE => {
                    if self.theme.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::theme::THEME_RECORD_TYPE,
                            message: "workbook contains more than one Theme record".to_string(),
                        });
                    }
                    let mut continues = Vec::new();
                    while records
                        .get(i + 1)
                        .is_some_and(|next| next.kind().get()
                            == crate::theme::CONTINUE_FRT12_RECORD_TYPE)
                    {
                        i += 1;
                        continues.push(records[i].payload().to_vec());
                    }
                    self.theme = Some(crate::theme::Theme::parse(record.payload(), &continues)?);
                },
                crate::style_ext::STYLE_EXT_RECORD_TYPE => {
                    self.style_extensions
                        .push(crate::style_ext::StyleExt::parse(record.payload())?);
                },
                crate::custom_view::USER_B_VIEW_RECORD_TYPE => {
                    self.custom_views
                        .push(crate::custom_view::WorkbookCustomView::parse(record.payload())?);
                },
                crate::real_time_data::REAL_TIME_DATA_RECORD_TYPE => {
                    // RTD = RealTimeData *ContinueFrt (MS-XLS 2.1): the
                    // logical payload is the record body plus any trailing
                    // ContinueFrt bodies.
                    let mut payload = record.payload().to_vec();
                    while records.get(i + 1).is_some_and(|next| {
                        next.kind().get()
                            == crate::real_time_data::CONTINUE_FRT_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(records[i].payload());
                    }
                    let previous_topic =
                        self.real_time_data.last().map(|topic| topic.topic.as_str());
                    self.real_time_data.push(
                        crate::real_time_data::RealTimeData::parse(&payload, previous_topic)?,
                    );
                },
                crate::web_pub::WEB_PUB_RECORD_TYPE => {
                    self.web_publications
                        .push(crate::web_pub::WebPub::parse(record.payload())?);
                },
                crate::mdx_metadata::MDT_INFO_RECORD_TYPE
                | crate::mdx_metadata::MDX_STR_RECORD_TYPE
                | crate::mdx_metadata::MDX_TUPLE_RECORD_TYPE
                | crate::mdx_metadata::MDX_SET_RECORD_TYPE
                | crate::mdx_metadata::MDX_PROP_RECORD_TYPE
                | crate::mdx_metadata::MDX_KPI_RECORD_TYPE
                | crate::mdx_metadata::MDB_RECORD_TYPE => {
                    // METADATA records are continued by ContinueFrt12 (MS-XLS
                    // 2.1): the logical payload is the record body plus any
                    // trailing ContinueFrt12 bodies.
                    let mut payload = record.payload().to_vec();
                    while records.get(i + 1).is_some_and(|next| {
                        next.kind().get()
                            == crate::mdx_metadata::CONTINUE_FRT12_RECORD_TYPE
                    }) {
                        i += 1;
                        payload.extend_from_slice(records[i].payload());
                    }
                    self.mdx_metadata
                        .push_record(record.kind().get(), &payload)?;
                },
                crate::book_ext::BOOK_EXT_RECORD_TYPE => {
                    if self.book_ext.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::book_ext::BOOK_EXT_RECORD_TYPE,
                            message: "workbook contains more than one BookExt record".to_string(),
                        });
                    }
                    self.book_ext = Some(crate::book_ext::BookExt::parse(record.payload())?);
                },
                crate::picture_compression::RECORD_TYPE => {
                    if self.picture_compression.is_some() {
                        return Err(Error::InvalidRecord {
                            record_type: crate::picture_compression::RECORD_TYPE,
                            message: "workbook contains more than one CompressPictures record"
                                .to_string(),
                        });
                    }
                    self.picture_compression = Some(
                        crate::picture_compression::Settings::parse(record.payload())?,
                    );
                },
                0x0809 => {
                    // BOF
                    let bof = BofRecord::parse(record.payload())?;
                    self.biff_version = bof.version;
                    self.is_1904_date_system = bof.is_1904_date_system;
                },
                0x0042
                    // CodePage
                    if record.payload().len() >= 2 => {
                        let codepage = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                        *encoding = Encoding::from_codepage(codepage)?;
                    },
                0x0022
                    // Date1904
                    if record.payload().len() >= 2 => {
                        let flag = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                        self.is_1904_date_system = flag == 1;
                    },
                0x0085 => {
                    // BoundSheet8
                    let sheet = BoundSheetRecord::parse(record.payload(), encoding)?;
                    bound_sheets.push(sheet);
                },
                0x00D5 => {
                    if record.payload().len() != 2 {
                        return Err(Error::InvalidLength { expected: 2, found: record.payload().len() });
                    }
                    let stream_id = litchi_core::binary::read_u16_le_at(record.payload(), 0)?;
                    if stream_id == 0 || self.pivot_cache_stream_ids.contains(&stream_id) {
                        return Err(Error::InvalidRecord {
                            record_type: 0x00D5,
                            message: "SXStreamID must be nonzero and unique".to_string(),
                        });
                    }
                    self.pivot_cache_stream_ids.push(stream_id);
                },
                0x01AE => {
                    // SUPBOOK: retain its position and whether it references
                    // this workbook so EXTERNSHEET indices remain stable.
                    self.formula_context.add_sup_book(record.payload());
                },
                0x0017 => {
                    // EXTERNSHEET: XTI entries used by PtgRef3d/PtgArea3d.
                    self.formula_context
                        .add_extern_sheet(record.payload())
                        .map_err(|message| Error::InvalidRecord {
                            record_type: 0x0017,
                            message: message.to_string(),
                        })?;
                },
                LBL_RECORD_TYPE => {
                    if defined_name_slots.len() == usize::from(u16::MAX) {
                        return Err(Error::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "workbook contains more than 65535 Lbl records".to_string(),
                        });
                    }
                    let record_index = u32::try_from(defined_name_slots.len() + 1)
                        .map_err(|_| Error::InvalidRecord {
                            record_type: LBL_RECORD_TYPE,
                            message: "Lbl record index overflows".to_string(),
                        })?;
                    let mut combined = record.payload().to_vec();
                    let mut continuation_chunks = Vec::new();
                    let mut next = i + 1;
                    while next < records.len() && records[next].kind().get() == 0x003c {
                        if combined
                            .len()
                            .checked_add(records[next].payload().len())
                            .is_none_or(|len| len > 1_048_576)
                        {
                            return Err(Error::InvalidRecord {
                                record_type: LBL_RECORD_TYPE,
                                message: "Lbl continuation data exceeds resource bound".to_string(),
                            });
                        }
                        continuation_chunks.push(records[next].payload().to_vec());
                        combined.extend_from_slice(records[next].payload());
                        next += 1;
                    }
                    defined_name_slots.push(DefinedNameSlot::parse_with_continuations(
                        &combined, record_index, continuation_chunks,
                    )?);
                    name_optional_target=Some((defined_name_slots.len()-1,0));
                    i=next-1;
                },
                NAME_CMT_RECORD_TYPE => {
                    let (target,stage)=name_optional_target.ok_or_else(|| Error::InvalidRecord {
                        record_type: NAME_CMT_RECORD_TYPE,
                        message: "NameCmt does not immediately follow a Lbl record".to_string(),
                    })?;
                    if stage!=0{return Err(Error::InvalidRecord{record_type:NAME_CMT_RECORD_TYPE,message:"NameCmt is duplicated or out of order in the Lbl optional-record sequence".to_string()})}
                    defined_name_slots[target].attach_comment(record.payload())?;
                    name_optional_target=Some((target,1));
                },
                crate::defined_names::NAME_FN_GRP12_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||Error::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 does not follow a Lbl record".to_string()})?;if stage>1{return Err(Error::InvalidRecord{record_type:crate::defined_names::NAME_FN_GRP12_RECORD_TYPE,message:"NameFnGrp12 is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_function_group(record.payload())?;name_optional_target=Some((target,2));},
                crate::defined_names::NAME_PUBLISH_RECORD_TYPE=>{let(target,stage)=name_optional_target.ok_or_else(||Error::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish does not follow a Lbl record".to_string()})?;if stage>2{return Err(Error::InvalidRecord{record_type:crate::defined_names::NAME_PUBLISH_RECORD_TYPE,message:"NamePublish is duplicated or out of order in the Lbl optional-record sequence".to_string()})}defined_name_slots[target].attach_publication(record.payload())?;name_optional_target=Some((target,3));},
                0x00FC => {
                    // SST
                    // SST may span multiple records, collect them all
                    let mut sst_records = vec![*record];
                    let mut sst_idx = i + 1;

                    // Collect all following CONTINUE records
                    while sst_idx < records.len() && records[sst_idx].kind().get() == 0x003C {
                        sst_records.push(records[sst_idx]);
                        sst_idx += 1;
                    }

                    let sst = SharedStringTable::parse_from_records(&sst_records, encoding)?;
                    self.shared_string_reference_count = sst.total_count;
                    strings.extend(sst.strings);
                    string_properties.extend(sst.properties);

                    // Skip the CONTINUE records we consumed
                    i = sst_idx - 1;
                },
                0x000A => {
                    // EOF - End of workbook globals
                    crate::font::validate_font_table(&self.fonts)?;
                    self.protection = protection_collector.finish()?;
                    self.calculation = calculation_collector.finish();
                    self.vba_metadata = vba_collector.finish();
                    self.environment = environment_collector.finish()?;
                    self.write_access = write_access_collector.finish();
                    self.table_styles = table_styles_collector
                        .finish(self.formatting.differential_formats().len())?;
                    self.shared_string_index = shared_string_index_collector.finish();
                    self.workbook_view = workbook_view_collector.finish(bound_sheets.len())?;
                    self.function_groups = function_group_collector.finish()?;
                    let extended_count=self.function_groups.as_ref().map_or(0,|groups|groups.extended_categories().len());for slot in defined_name_slots.iter(){slot.validate_extended_category(extended_count)?;}
                    self.external_links = external_link_collector.finish(bound_sheets.len())?;
                    break;
                },
                _ => {
                    // Skip other records for now
                },
            }
            i += 1;
        }

        Ok(())
    }
}
