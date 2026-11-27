/// Passive document-level privacy-removal requests.
///
/// These flags are retained for round trips only. This crate does not remove,
/// anonymize, rewrite, or suppress document properties, comments, or times.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct DocumentPrivacyPolicies {
    /// `\rempersonalinfo`: an emitting application is asked to remove personal
    /// information such as author names in properties or comments.
    pub remove_personal_information: bool,
    /// `\remdttm`: an emitting application is asked to remove date/time
    /// information in properties or comments.
    pub remove_date_time_information: bool,
}

impl DocumentPrivacyPolicies {
    /// Return whether no privacy-removal request was present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        !self.remove_personal_information && !self.remove_date_time_information
    }
}
