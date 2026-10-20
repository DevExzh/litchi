//! `PowerPoint` 9 picture-bullet collections and their package-level edits.
//!
//! [`crate::Package::edit_picture_bullets`] stages lifecycle edits of the
//! live document's picture-bullet collection; see [`Edit`].

mod edit;
mod record;

pub use edit::Edit;
pub use record::{
    Change, ChangeKind, MAX_PICTURE_BULLET_INDEX, PackageLimits, PictureBullet,
    PictureBulletChange, PictureBulletChangeKind, PictureBulletCollection, PictureBulletLimits,
    PictureBulletPackageLimits, PictureBulletRevision, PictureBulletSnapshot, PictureBulletType,
    Revision, Snapshot,
};

crate::scoped::family!("picture-bullet");
