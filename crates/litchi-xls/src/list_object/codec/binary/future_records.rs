//! Retained BIFF future-record state shared by the list-object codecs.

/// A Feature11/Feature12 record together with its continuation fragments.
///
/// The package collector owns the assembly policy; this wire type only keeps
/// the exact base, continuation, and combined payloads available to the
/// semantic decoders and lossless writers.
pub(in crate::list_object) struct PendingFeature {
    pub(in crate::list_object) record_type: u16,
    pub(in crate::list_object) base: Vec<u8>,
    pub(in crate::list_object) continuations: Vec<Vec<u8>>,
    pub(in crate::list_object) combined: Vec<u8>,
}
