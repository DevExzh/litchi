//! Lossless storage for unsupported or future list-object records.

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueListObjectFeature {
    pub(in crate::list_object) record_type: u16,
    pub(in crate::list_object) base_payload: Vec<u8>,
    pub(in crate::list_object) continuation_payloads: Vec<Vec<u8>>,
}
impl OpaqueListObjectFeature {
    #[must_use]
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }
    #[must_use]
    pub fn base_payload(&self) -> &[u8] {
        &self.base_payload
    }
    #[must_use]
    pub fn continuation_payloads(&self) -> &[Vec<u8>] {
        &self.continuation_payloads
    }
    #[must_use]
    pub fn total_payload_len(&self) -> usize {
        self.base_payload.len()
            + self
                .continuation_payloads
                .iter()
                .map(|v| v.len() - 12)
                .sum::<usize>()
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct OpaqueListObjectFutureRecord {
    pub(in crate::list_object) record_type: u16,
    pub(in crate::list_object) payload: Vec<u8>,
    pub(in crate::list_object) continuation_payloads: Vec<Vec<u8>>,
    pub(in crate::list_object) after_list12_count: usize,
}
impl OpaqueListObjectFutureRecord {
    #[must_use]
    pub const fn record_type(&self) -> u16 {
        self.record_type
    }
    #[must_use]
    pub fn payload(&self) -> &[u8] {
        &self.payload
    }
    #[must_use]
    pub fn continuation_payloads(&self) -> &[Vec<u8>] {
        &self.continuation_payloads
    }
    #[must_use]
    pub const fn after_list12_count(&self) -> usize {
        self.after_list12_count
    }
}
