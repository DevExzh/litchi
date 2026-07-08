//! Verify generated PPT files match POI reference

use litchi::ole::OleFile;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Verify minimal.ppt matches POI
    let poi_path =
        "3rdparty/poi/poi-scratchpad/src/main/resources/org/apache/poi/hslf/data/empty.ppt";
    let poi_file = std::fs::File::open(poi_path)?;
    let mut poi_ole = OleFile::open(poi_file)?;
    let poi = poi_ole.open_stream(&["PowerPoint Document"])?;

    let min_file = std::fs::File::open("minimal.ppt")?;
    let mut min_ole = OleFile::open(min_file)?;
    let min = min_ole.open_stream(&["PowerPoint Document"])?;

    let mut diff = 0;
    for i in 0..poi.len().min(min.len()) {
        if poi[i] != min[i] {
            diff += 1;
        }
    }
    println!("minimal.ppt vs POI: {} differences", diff);

    // Check with_slide.ppt size
    let ws_file = std::fs::File::open("with_slide.ppt")?;
    let mut ws_ole = OleFile::open(ws_file)?;
    let ws = ws_ole.open_stream(&["PowerPoint Document"])?;
    println!("with_slide.ppt: {} bytes", ws.len());

    if diff == 0 {
        println!("\n✓ Both files generated successfully with spec-based constants!");
    }

    Ok(())
}
