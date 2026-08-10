#![forbid(unsafe_code)]

//! Raw generated Protocol Buffer types used by Apple iWork IWA archives.
//!
//! This crate owns only the schema and code-generation boundary. It does not
//! decode IWA objects, traverse packages, or provide application-specific
//! semantics. Consumers that need those behaviors should depend on the
//! appropriate `litchi-pages`, `litchi-keynote`, `litchi-numbers`, or
//! `litchi-iwa-archive` crate instead.

/// Generated source is kept behind one audited boundary so workspace lints
/// continue to apply to every hand-written item in this crate.
#[doc(hidden)]
mod generated {
    #![allow(
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "prost-build output is generated from the native IWA schemas."
    )]

    include!(concat!(env!("OUT_DIR"), "/iwa_protos.rs"));
}

/// Private Buffa eager/lazy view sidecar for archive adapters.
///
/// The sidecar is deliberately not an untrusted-ingress API. In Buffa 0.9.1,
/// deferred lazy message access rebuilds its decode context with recursion and
/// unknown-field limits only; it does not retain the original
/// `DecodeOptions::with_element_memory_limit` budget. Nested deferred
/// allocations can therefore escape that initial element-memory accounting.
/// Archive adapters must establish their own complete resource policy before
/// accepting untrusted payloads through this path.
#[doc(hidden)]
mod buffa_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa sidecar is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the native IWA schemas."
    )]

    include!(concat!(env!("OUT_DIR"), "/buffa/iwa_buffa_protos.rs"));
}

/// Private Buffa lazy-view projection for `TSWP.StorageArchive.text` only.
///
/// It is generated in isolation with unknown retention disabled. The source
/// IWA payload remains the byte-authoritative preservation representation.
#[doc(hidden)]
mod buffa_text_storage_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-text-storage/iwa_text_storage_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the GroupNode category-label path.
///
/// It includes only an empty node envelope, UUID identity, and the scalar
/// Boolean, Date, Number, and String wrappers. The adapter streams child and
/// CellValue routing from source bytes, which remain authoritative for
/// preservation.
#[doc(hidden)]
mod buffa_group_node_category_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-group-node-category/iwa_group_node_category_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for the Keynote root show reference.
///
/// The required root base archive remains opaque. Only the nested show
/// identifier is decoded, while caller-owned IWA bytes remain authoritative.
#[doc(hidden)]
mod buffa_keynote_document_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-document/iwa_keynote_document_buffa_protos.rs"
    ));
}

/// Private strict Buffa lazy-view projection for Keynote placeholder text.
///
/// Generated code sees only the singular inheritance chain, optional kind,
/// and optional owned-storage edge. Caller-owned bytes remain authoritative.
#[doc(hidden)]
mod buffa_keynote_placeholder_text_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-placeholder-text/iwa_keynote_placeholder_text_buffa_protos.rs"
    ));
}

/// Private strict Buffa lazy-view projection for focused Keynote slide owners.
///
/// Generated code sees only exact references, selector-facing slide scalars,
/// semantic note/title/body edges, and required slide/note envelopes. Unknown
/// bytes remain solely in caller-owned IWA and are never retained here.
#[doc(hidden)]
mod buffa_keynote_speaker_notes_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-speaker-notes/iwa_keynote_speaker_notes_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for `TST.TableInfoArchive.table_model`.
///
/// The required drawable base archive and all unselected table metadata remain
/// caller-owned opaque source bytes. Generated code sees only the model
/// reference identifier after the Numbers adapter's strict raw preflight.
#[doc(hidden)]
mod buffa_table_info_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-table-info/iwa_table_info_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Keynote show settings.
///
/// The repeated slide tree is deliberately absent from generated code. A
/// bounded handwritten router streams its references directly from the
/// caller-owned payload, while this projection validates direct references,
/// size, and scalar show settings.
#[doc(hidden)]
mod buffa_keynote_show_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-show/iwa_keynote_show_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Pages section pagination.
///
/// Only the three scalar pagination fields are generated. Caller-owned section
/// bytes remain authoritative for preservation and rewriting.
#[doc(hidden)]
mod buffa_pages_section_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-section/iwa_pages_section_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Pages root/body graph leaves.
///
/// It contains two root references and one singular section-boundary entry.
/// The repeated enclosing table and opaque base document stay in caller-owned
/// source bytes and never enter generated code.
#[doc(hidden)]
mod buffa_pages_body_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-pages-body/iwa_pages_body_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Numbers sheet, form, and table names.
#[doc(hidden)]
mod buffa_numbers_names_generated {
    #![allow(
        elided_lifetimes_in_paths,
        unreachable_pub,
        clippy::all,
        clippy::allow_attributes_without_reason,
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        non_snake_case,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "Buffa generated projection is private to this crate."
    )]
    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-numbers-names/iwa_numbers_names_buffa_protos.rs"
    ));
}

/// Private Buffa lazy-view projection for Keynote slide transitions.
///
/// It contains only the nested transition attributes and slide-node transition
/// flag.  The format crate owns strict ingress validation and retains the raw
/// IWA source for every write.
#[doc(hidden)]
mod buffa_keynote_slide_transition_generated {
    #![allow(
        elided_lifetimes_in_paths,
        reason = "Buffa 0.9.1 generated views elide explicit lifetimes."
    )]
    #![allow(
        unreachable_pub,
        reason = "The Buffa projection is intentionally private to this crate."
    )]
    #![allow(
        clippy::allow_attributes_without_reason,
        reason = "Buffa 0.9.1 generated source contains internal lint allowances."
    )]
    #![allow(
        clippy::map_err_ignore,
        clippy::shadow_reuse,
        clippy::shadow_same,
        reason = "Buffa 0.9.1 generated decoders use these implementation patterns."
    )]
    #![allow(
        non_snake_case,
        clippy::all,
        clippy::arbitrary_source_item_ordering,
        clippy::module_name_repetitions,
        clippy::pedantic,
        reason = "buffa-build output is generated from the derived wire projection."
    )]

    include!(concat!(
        env!("OUT_DIR"),
        "/buffa-keynote-slide-transition/iwa_keynote_slide_transition_buffa_protos.rs"
    ));
}

/// Internal archive-header codec implemented by the private Buffa sidecar.
///
/// This module intentionally exchanges only the existing Prost compatibility
/// structs. Buffa-generated types remain an implementation detail and cannot
/// leak into the public format APIs.
#[doc(hidden)]
pub mod archive_codec;

/// Internal raw-text projection codec implemented by the private Buffa
/// sidecar. Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod text_storage_codec;

/// Internal raw GroupNode category-label projection implemented by the
/// private Buffa sidecar. Generated types remain inaccessible to downstream
/// crates.
#[doc(hidden)]
pub mod group_node_category_codec;

/// Internal Keynote root-document projection implemented by the private
/// Buffa sidecar. Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod keynote_document_codec;

/// Internal strict Keynote placeholder text-owner projection. Generated types
/// remain inaccessible and source bytes stay authoritative.
#[doc(hidden)]
pub mod keynote_placeholder_text_codec;

/// Internal strict Keynote speaker-note owner projection. Generated types
/// remain inaccessible to downstream crates and source bytes stay authoritative.
#[doc(hidden)]
pub mod keynote_speaker_notes_codec;

/// Internal Numbers TableInfo model-reference projection implemented by a
/// private strict Buffa lazy-view sidecar. Generated types remain inaccessible
/// to downstream crates.
#[doc(hidden)]
pub mod table_info_codec;

/// Internal Keynote show projection and bounded streaming slide-tree codec.
/// Generated types remain inaccessible to downstream crates.
#[doc(hidden)]
pub mod keynote_show_codec;

/// Internal Pages section-pagination projection implemented by a private
/// Buffa lazy-view sidecar. Generated types remain inaccessible downstream.
#[doc(hidden)]
pub mod pages_section_codec;

/// Internal Pages root/body projection implemented by a private strict Buffa
/// lazy-view adapter. Generated types remain inaccessible downstream.
#[doc(hidden)]
pub mod pages_body_codec;

/// Internal strict Pages document page-layout projection. Generated types stay
/// private and caller-owned raw bytes remain authoritative.
#[doc(hidden)]
pub mod pages_page_layout_codec;

/// Internal strict Pages document-settings projection. Generated types remain
/// private and caller-owned raw bytes stay authoritative.
#[doc(hidden)]
pub mod pages_document_settings_codec;

/// Internal strict Numbers names projection. Generated types remain private
/// and all decoded names borrow caller-owned source bytes.
#[doc(hidden)]
pub mod numbers_names_codec;

/// Internal Keynote slide-transition projection implemented by a private
/// Buffa lazy-view sidecar. Generated types remain inaccessible downstream.
#[doc(hidden)]
pub mod keynote_slide_transition_codec;

pub use generated::{
    kn, knsos, tn, tnsos, tp, tpsos, tsa, tsasos, tsce, tsch, tschsos, tsck, tscksos, tsd, tsdsos,
    tsk, tsp, tss, tsssos, tst, tstsos, tswp, tswpsos,
};

#[cfg(test)]
mod tests {
    use buffa::{LazyMessageView as _, Message as _};
    use prost::Message as _;

    #[test]
    fn generated_messages_round_trip_without_runtime_names() -> Result<(), prost::DecodeError> {
        let input = super::tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: Vec::new(),
            should_merge: Some(true),
        };
        let encoded = input.encode_to_vec();
        let decoded = super::tsp::ArchiveInfo::decode(encoded.as_slice())?;
        assert_eq!(decoded, input);
        Ok(())
    }

    #[test]
    fn buffa_archive_info_matches_prost_wire_format() -> Result<(), Box<dyn std::error::Error>> {
        let input = super::tsp::ArchiveInfo {
            identifier: Some(42),
            message_infos: Vec::new(),
            should_merge: Some(true),
        };
        let prost_encoded = input.encode_to_vec();

        let buffa_decoded =
            super::buffa_generated::TSP::ArchiveInfo::decode_from_slice(&prost_encoded)?;
        assert_eq!(buffa_decoded.identifier, input.identifier);
        assert!(buffa_decoded.message_infos.is_empty());
        assert_eq!(buffa_decoded.should_merge, input.should_merge);

        let buffa_encoded = buffa_decoded.try_encode_to_vec()?;
        let prost_decoded = super::tsp::ArchiveInfo::decode(buffa_encoded.as_slice())?;
        assert_eq!(prost_decoded, input);
        Ok(())
    }

    #[test]
    fn buffa_archive_info_lazy_view_round_trips() -> Result<(), Box<dyn std::error::Error>> {
        let input = super::buffa_generated::TSP::ArchiveInfo {
            identifier: Some(42),
            message_infos: vec![super::buffa_generated::TSP::MessageInfo {
                r#type: 7,
                length: 11,
                ..Default::default()
            }],
            should_merge: Some(true),
            ..Default::default()
        };
        let encoded = input.try_encode_to_vec()?;
        let lazy: super::buffa_generated::TSP::ArchiveInfoLazyView<'_> =
            buffa::DecodeOptions::new().decode_lazy_view(&encoded)?;

        assert_eq!(lazy.message_infos.len(), 1);
        let message_info_view = lazy.message_infos.try_get(0)?;
        assert_eq!(
            message_info_view.map(|view| (view.r#type, view.length)),
            Some((7, 11))
        );
        assert_eq!(lazy.to_owned_message()?, input);
        assert_eq!(lazy.try_encode_to_vec()?, encoded);
        Ok(())
    }
}
