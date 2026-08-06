//! Embedded media and package serialization semantics.

#![allow(unused_imports)]

use super::*;
use crate::media::MediaAssetId;

impl NumbersEditor {
    /// List metadata-backed media reachable from this spreadsheet package.
    pub fn media_assets(&self) -> Result<Vec<EmbeddedMediaAsset>> {
        reachable_embedded_assets(&self.package, [1])
    }

    /// List media reachable from one sheet and its drawable object graph.
    pub fn sheet_media_assets(&self, sheet_id: u64) -> Result<Vec<EmbeddedMediaAsset>> {
        if !self
            .sheets()?
            .iter()
            .any(|sheet| sheet.object_id == sheet_id)
        {
            return Err(Error::ParseError(format!(
                "Numbers sheet object {sheet_id} is not reachable"
            )));
        }
        reachable_embedded_assets(&self.package, [sheet_id])
    }

    pub fn extract_media(&self, data_identifier: u64) -> Result<Vec<u8>> {
        let data_identifier = MediaAssetId::try_from(data_identifier)?;
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Numbers object graph"
            )));
        }
        IWorkMediaEditor::from_package(self.package.clone())?.extract(data_identifier)
    }

    /// Replace a referenced materialized asset without changing its data identifier.
    pub fn replace_media(&mut self, data_identifier: u64, replacement: &[u8]) -> Result<Vec<u8>> {
        let data_identifier = MediaAssetId::try_from(data_identifier)?;
        if !self
            .media_assets()?
            .iter()
            .any(|asset| asset.data_identifier == data_identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Data identifier {data_identifier} is not reachable from the Numbers object graph"
            )));
        }
        let mut media = IWorkMediaEditor::from_package(self.package.clone())?;
        let old = media.replace(data_identifier, replacement)?;
        let staged = media.into_package();
        Self::from_package(staged.clone())?;
        self.package = staged;
        Ok(old)
    }

    pub fn into_package(self) -> IWorkPackage {
        self.package
    }

    pub fn to_bytes(&self) -> Result<Vec<u8>> {
        self.package.to_bytes()
    }

    pub fn save(&self, path: impl AsRef<Path>) -> Result<()> {
        self.package.save(path)
    }
}
