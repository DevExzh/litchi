#![forbid(unsafe_code)]

//! Raw generated Protocol Buffer types used by Apple iWork IWA archives.
//!
//! This crate owns only the schema and code-generation boundary. It does not
//! decode IWA objects, traverse packages, or provide application-specific
//! semantics. Consumers that need those behaviors should depend on
//! [`litchi_iwa`](https://docs.rs/litchi-iwa) instead.

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

pub use generated::{
    kn, knsos, tn, tnsos, tp, tpsos, tsa, tsasos, tsce, tsch, tschsos, tsck, tscksos, tsd, tsdsos,
    tsk, tsp, tss, tsssos, tst, tstsos, tswp, tswpsos,
};

#[cfg(test)]
mod tests {
    use prost::{Message, Name};

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
    fn generated_schema_names_remain_available() {
        assert_eq!(super::tsp::ArchiveInfo::full_name(), "TSP.ArchiveInfo");
        assert_eq!(
            super::tp::DocumentArchive::full_name(),
            "TP.DocumentArchive"
        );
        assert_eq!(
            super::tsch::pre_uff::ChartInfoArchive::full_name(),
            "TSCH.PreUFF.ChartInfoArchive"
        );
    }
}
