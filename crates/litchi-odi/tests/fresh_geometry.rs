//! Geometry inserted on a fresh image whose compact `content.xml` dropped
//! the unused `svg` namespace declaration (record 0288, 0282 G4).

use litchi_odi::Package;

fn geometry(package: &Package) -> [Option<String>; 4] {
    let image = package.image();
    let frame = &image.frames()[0];
    [
        frame.x().map(str::to_owned),
        frame.y().map(str::to_owned),
        frame.width().map(str::to_owned),
        frame.height().map(str::to_owned),
    ]
}

#[test]
fn fresh_image_geometry_round_trips() {
    let package = Package::new().unwrap();
    let mut edit = package.edit();
    edit.set_geometry(
        0,
        Some("1cm".into()),
        Some("2cm".into()),
        Some("3cm".into()),
        Some("4cm".into()),
    )
    .unwrap();
    let next = edit.commit().unwrap().into_package();
    assert_eq!(
        geometry(&next),
        [
            Some("1cm".into()),
            Some("2cm".into()),
            Some("3cm".into()),
            Some("4cm".into())
        ]
    );
    let reopened = Package::from_bytes(next.as_bytes()).unwrap();
    assert_eq!(geometry(&reopened), geometry(&next));
    // A second geometry edit reuses the declaration.
    let mut edit = next.edit();
    edit.set_x(0, Some("5cm".into())).unwrap();
    let again = edit.commit().unwrap().into_package();
    assert_eq!(geometry(&again)[0].as_deref(), Some("5cm"));
}

#[test]
fn fresh_image_title_and_description_round_trip() {
    let package = Package::new().unwrap();
    let mut edit = package.edit();
    edit.set_frame_title(0, Some("T".into())).unwrap();
    edit.set_frame_description(0, Some("D".into())).unwrap();
    let next = edit.commit().unwrap().into_package();
    let image = next.image();
    assert_eq!(image.frames()[0].title(), Some("T"));
    assert_eq!(image.frames()[0].description(), Some("D"));
}
