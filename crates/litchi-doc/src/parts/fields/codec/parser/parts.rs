//! Typed intermediate parts shared by semantic field parsers.

use crate::parts::fields::model::*;

pub(in crate::parts::fields) struct DdeParts {
    pub(in crate::parts::fields) kind: DdeFieldKind,
    pub(in crate::parts::fields) application: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) representation: Option<DdeRepresentation>,
    pub(in crate::parts::fields) omit_graphic_data: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct HyperlinkParts {
    pub(in crate::parts::fields) external_target: Option<String>,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) screen_tip: Option<String>,
    pub(in crate::parts::fields) target_frame: Option<String>,
    pub(in crate::parts::fields) appends_image_map_coordinates: bool,
    pub(in crate::parts::fields) opens_new_window: bool,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct LinkParts {
    pub(in crate::parts::fields) application_type: String,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) item: Option<String>,
    pub(in crate::parts::fields) automatic_updates: bool,
    pub(in crate::parts::fields) result_options: Vec<LinkResultOption>,
    pub(in crate::parts::fields) formatting_modes: Vec<LinkFormatting>,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ExternalIncludeParts {
    pub(in crate::parts::fields) kind: IncludeFieldKind,
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) bookmark: Option<String>,
    pub(in crate::parts::fields) suppress_nested_field_updates: bool,
    pub(in crate::parts::fields) omit_picture_data: bool,
    pub(in crate::parts::fields) options: Vec<ExternalIncludeOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfContentsParts {
    pub(in crate::parts::fields) options: Vec<TableOfContentsOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfContentsEntryParts {
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<TableOfContentsEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfAuthoritiesEntryParts {
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct IndexEntryParts {
    pub(in crate::parts::fields) entry: String,
    pub(in crate::parts::fields) options: Vec<IndexEntryOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ReferencedDocumentParts {
    pub(in crate::parts::fields) source: String,
    pub(in crate::parts::fields) relative_path: bool,
    pub(in crate::parts::fields) switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct TableOfAuthoritiesParts {
    pub(in crate::parts::fields) options: Vec<TableOfAuthoritiesOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct IndexParts {
    pub(in crate::parts::fields) options: Vec<IndexOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct ReferenceParts {
    pub(in crate::parts::fields) bookmark: String,
    pub(in crate::parts::fields) options: Vec<ReferenceFieldOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct StyleReferenceParts {
    pub(in crate::parts::fields) style_name: String,
    pub(in crate::parts::fields) options: Vec<StyleReferenceFieldOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct AutoTextParts {
    pub(in crate::parts::fields) entry_name: String,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}

pub(in crate::parts::fields) struct AutoTextListParts {
    pub(in crate::parts::fields) display_text: Option<String>,
    pub(in crate::parts::fields) options: Vec<AutoTextListOption>,
    pub(in crate::parts::fields) unknown_switches: Vec<MergeFieldSwitch>,
}
