//! Debug PPTX - test features one by one
//!
//! Run with: cargo run --example pptx_debug

use litchi::ooxml::pptx::Package;
use std::fs;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // Test 1: Media on multiple slides
    println!("Test 1: Multiple slides with media...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Slide 1 - Audio");
        let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
        slide1.add_audio(audio_data, 914400, 1828800, 914400, 914400);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Slide 2 - Video");
        let video_data = fs::read("ForBiggerMeltdowns.mp4")?;
        slide2.add_video(video_data, 914400, 1828800, 4572000, 2571750);

        pkg.save("test1_multi_media.pptx")?;
        println!("  -> test1_multi_media.pptx");
    }

    // Test 2: Media + Table
    println!("Test 2: Media + Table...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Audio");
        let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
        slide1.add_audio(audio_data, 914400, 1828800, 914400, 914400);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Table");
        let data = vec![
            vec!["A".to_string(), "B".to_string()],
            vec!["1".to_string(), "2".to_string()],
        ];
        slide2.add_table(data, 914400, 1828800, 5486400, 1828800);

        pkg.save("test2_media_table.pptx")?;
        println!("  -> test2_media_table.pptx");
    }

    // Test 3: Media + Comments
    println!("Test 3: Media + Comments...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Audio");
        let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
        slide1.add_audio(audio_data, 914400, 1828800, 914400, 914400);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Comments");
        slide2.add_comment(0, "Test comment", 914400, 914400);

        pkg.save("test3_media_comments.pptx")?;
        println!("  -> test3_media_comments.pptx");
    }

    // Test 4: Two audio files on one slide
    println!("Test 4: Two audio on same slide...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Two Audio Files");
        let mp3_data = fs::read("file_example_MP3_700KB.mp3")?;
        slide1.add_audio(mp3_data, 914400, 1828800, 914400, 914400);
        let wav_data = fs::read("file_example_WAV_1MG.wav")?;
        slide1.add_audio(wav_data, 2743200, 1828800, 914400, 914400);

        pkg.save("test4_two_audio.pptx")?;
        println!("  -> test4_two_audio.pptx");
    }

    // Test 5: Media + Group shapes
    println!("Test 5: Media + Group shapes...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide1 = pres.add_slide()?;
        slide1.set_title("Audio");
        let audio_data = fs::read("file_example_MP3_700KB.mp3")?;
        slide1.add_audio(audio_data, 914400, 1828800, 914400, 914400);

        let slide2 = pres.add_slide()?;
        slide2.set_title("Groups");
        let group_idx = slide2.add_group(914400, 1828800, 3657600, 2743200);
        slide2.add_rectangle_to_group(group_idx, 0, 0, 914400, 914400, Some("FF0000".to_string()));

        pkg.save("test5_media_groups.pptx")?;
        println!("  -> test5_media_groups.pptx");
    }

    // Test 6: Comments only (no media)
    println!("Test 6: Comments only...");
    {
        let mut pkg = Package::new()?;
        let pres = pkg.presentation_mut()?;

        let slide = pres.add_slide()?;
        slide.set_title("Comments Only");
        slide.add_comment(0, "Test comment", 914400, 914400);

        pkg.save("test6_comments_only.pptx")?;
        println!("  -> test6_comments_only.pptx");
    }

    println!("\nDone! Test each file in PowerPoint to find the problematic combination.");

    Ok(())
}
