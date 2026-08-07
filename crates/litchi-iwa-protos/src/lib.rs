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
