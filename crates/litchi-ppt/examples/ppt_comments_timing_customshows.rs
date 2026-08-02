// Example: Generate PPT files demonstrating comments, slide timings, and custom shows.
// Open the generated files in PowerPoint to verify each feature.

use litchi_ppt::PptWriter;
use litchi_ppt::writer::comments::{CommentDateTime, SlideComment};
use litchi_ppt::writer::custom_shows::CustomShow;
use litchi_ppt::writer::slide_timing::SlideTiming;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // ── 1. Comments ────────────────────────────────────────────────────
    // In PowerPoint: Review tab → Show Comments should display these.
    {
        let mut w = PptWriter::new();

        let s0 = w.add_slide()?;
        w.add_textbox(s0, 50, 50, 400, 40, "Slide 1 – has two comments")?;
        w.add_comment(
            s0,
            SlideComment::new("Alice Smith", "This slide needs a better title.", 72, 36).with_date(
                CommentDateTime {
                    year: 2025,
                    month: 6,
                    day: 15,
                    hour: 10,
                    minute: 30,
                    second: 0,
                    millisecond: 0,
                },
            ),
        )?;
        w.add_comment(
            s0,
            SlideComment::new("Bob Jones", "Agreed, let's revise.", 200, 100).with_date(
                CommentDateTime {
                    year: 2025,
                    month: 6,
                    day: 15,
                    hour: 11,
                    minute: 5,
                    second: 0,
                    millisecond: 0,
                },
            ),
        )?;

        let s1 = w.add_slide()?;
        w.add_textbox(s1, 50, 50, 400, 40, "Slide 2 – one comment")?;
        w.add_comment(
            s1,
            SlideComment::new("Charlie Brown", "Nice chart on this slide.", 150, 200),
        )?;

        w.save("ppt_comments.ppt")?;
        println!("Created ppt_comments.ppt  (2 slides, 3 comments total)");
    }

    // ── 2. Slide timings ───────────────────────────────────────────────
    // In PowerPoint: Slide Show → Set Up Show / Rehearse Timings, or
    // simply run the slide show to observe auto-advance behaviour.
    {
        let mut w = PptWriter::new();

        let s0 = w.add_slide()?;
        w.add_textbox(
            s0,
            50,
            50,
            500,
            40,
            "Slide 1 – auto-advances after 2 seconds",
        )?;
        w.set_slide_timing(s0, SlideTiming::auto_advance(2000))?;

        let s1 = w.add_slide()?;
        w.add_textbox(
            s1,
            50,
            50,
            500,
            40,
            "Slide 2 – auto-advances after 5 seconds",
        )?;
        w.set_slide_timing(s1, SlideTiming::auto_advance(5000))?;

        let s2 = w.add_slide()?;
        w.add_textbox(s2, 50, 50, 500, 40, "Slide 3 – click only (no timer)")?;
        w.set_slide_timing(s2, SlideTiming::on_click_only())?;

        let s3 = w.add_slide()?;
        w.add_textbox(s3, 50, 50, 500, 40, "Slide 4 – HIDDEN slide")?;
        w.set_slide_timing(s3, SlideTiming::hidden())?;

        let s4 = w.add_slide()?;
        w.add_textbox(s4, 50, 50, 500, 40, "Slide 5 – auto 3s, no click advance")?;
        w.set_slide_timing(
            s4,
            SlideTiming::auto_advance(3000).with_click_advance(false),
        )?;

        w.save("ppt_timings.ppt")?;
        println!(
            "Created ppt_timings.ppt   (5 slides: 2s auto, 5s auto, click-only, hidden, 3s no-click)"
        );
    }

    // ── 3. Custom slide shows ──────────────────────────────────────────
    // In PowerPoint: Slide Show → Custom Slide Show should list two shows.
    {
        let mut w = PptWriter::new();

        let s0 = w.add_slide()?;
        w.add_textbox(s0, 50, 50, 500, 40, "Slide 1 – Introduction")?;

        let s1 = w.add_slide()?;
        w.add_textbox(s1, 50, 50, 500, 40, "Slide 2 – Technical Details")?;

        let s2 = w.add_slide()?;
        w.add_textbox(s2, 50, 50, 500, 40, "Slide 3 – Budget Overview")?;

        let s3 = w.add_slide()?;
        w.add_textbox(s3, 50, 50, 500, 40, "Slide 4 – Timeline")?;

        let s4 = w.add_slide()?;
        w.add_textbox(s4, 50, 50, 500, 40, "Slide 5 – Conclusion")?;

        // "Executive Summary" shows only intro, budget, and conclusion
        w.add_custom_show(CustomShow::new("Executive Summary", &[0, 2, 4]));

        // "Full Technical" shows everything except the budget slide
        w.add_custom_show(CustomShow::new("Full Technical", &[0, 1, 3, 4]));

        w.save("ppt_custom_shows.ppt")?;
        println!(
            "Created ppt_custom_shows.ppt (5 slides, 2 custom shows: 'Executive Summary' [1,3,5], 'Full Technical' [1,2,4,5])"
        );
    }

    // ── 4. Combined: all three features ────────────────────────────────
    {
        let mut w = PptWriter::new();

        let s0 = w.add_slide()?;
        w.add_textbox(s0, 50, 50, 500, 40, "Slide 1 – Welcome")?;
        w.set_slide_timing(s0, SlideTiming::auto_advance(3000))?;
        w.add_comment(
            s0,
            SlideComment::new("Reviewer", "Opening slide looks good.", 72, 36).with_date(
                CommentDateTime {
                    year: 2025,
                    month: 7,
                    day: 1,
                    hour: 9,
                    minute: 0,
                    ..Default::default()
                },
            ),
        )?;

        let s1 = w.add_slide()?;
        w.add_textbox(s1, 50, 50, 500, 40, "Slide 2 – Agenda")?;
        w.set_slide_timing(s1, SlideTiming::auto_advance(4000))?;

        let s2 = w.add_slide()?;
        w.add_textbox(s2, 50, 50, 500, 40, "Slide 3 – Deep Dive")?;
        w.set_slide_timing(s2, SlideTiming::on_click_only())?;
        w.add_comment(
            s2,
            SlideComment::new("Manager", "Add more data here.", 100, 150),
        )?;

        let s3 = w.add_slide()?;
        w.add_textbox(s3, 50, 50, 500, 40, "Slide 4 – Summary")?;
        w.set_slide_timing(s3, SlideTiming::auto_advance(2000))?;

        w.add_custom_show(CustomShow::new("Quick Overview", &[0, 3]));
        w.add_custom_show(CustomShow::new("Full Presentation", &[0, 1, 2, 3]));

        w.save("ppt_all_features.ppt")?;
        println!("Created ppt_all_features.ppt (4 slides, comments + timings + 2 custom shows)");
    }

    println!("\nAll files created. Open in PowerPoint to verify:");
    println!("  - ppt_comments.ppt       → Review > Show Comments");
    println!("  - ppt_timings.ppt        → Run Slide Show (observe auto-advance & hidden slide)");
    println!("  - ppt_custom_shows.ppt   → Slide Show > Custom Slide Show");
    println!("  - ppt_all_features.ppt   → All of the above combined");

    Ok(())
}
