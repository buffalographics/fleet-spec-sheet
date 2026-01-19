Fleet Spec Sheet Generator — Illustrator Script (V1)

Version: Illustrator V1
Status: Production-ready
Platform: Adobe Illustrator (ExtendScript / .jsx)

⸻

Overview

This Illustrator script generates Fleet Vehicle Spec Sheets directly inside the active Illustrator document, using existing artboards as the data source.

It eliminates file-path dependency and mirrors real production workflow by pulling:
• Proof imagery from a dedicated PROOF artboard
• Print items from individual artwork artboards
• Quantities from artboard names

The output is one or more SPEC SHEET (V1) artboards formatted for printing and pen sign-off.

⸻

What It Generates

Each spec sheet includes: 1. Header
• BUFFALO GRAPHICS COMPANY branding 2. Job Details
• Customer (prompted)
• Vehicle type (prompted)
• Vehicle Unit # (intentionally blank for handwriting) 3. Proof Section
• Automatically placed from the PROOF / MOCKUP artboard
• Scaled proportionally and centered
• Framed to match production spec layout 4. Print Manifest
• Thumbnail (auto-exported from each artboard)
• Name (artboard name)
• File Name (artboard name)
• Quantity (parsed automatically)
• Done column (blank) 5. Install Sign-Off
• Installed By
• Photos Taken (YES / NO)
• Date
• Issues / Damage Notes

Pagination is handled automatically if the manifest exceeds one page.

⸻

Artboard Rules (Critical)

Proof Artboard

Exactly one artboard must be named:
• PROOF or
• MOCKUP

(case-insensitive)

This artboard is used as the proof image.

⸻

Ignored Artboards

Any artboard with a name matching:
• unit number
• unit numbers
• unit #
• unit no

(case-insensitive)

These are skipped entirely.

⸻

Print Manifest Artboards

All other artboards are treated as print items.

Each artboard:
• Becomes one row in the Print Manifest
• Is exported automatically as a thumbnail

⸻

Quantity Parsing

Quantities are parsed from the artboard name.

Supported formats (case-insensitive):
• QTY2
• QTY_2
• QTY-2
• QTY 2
• x2
• 2x

If no quantity is found, default = 1.

Examples:
• DOOR_DRIVER_QTY_2 → Qty = 2
• USDOT x3 → Qty = 3
• LOGO_PASSENGER → Qty = 1

⸻

Font Handling
• Script attempts to use HudsonNY everywhere
• Searches installed fonts by:
• Exact name
• Family
• Style
• Partial match containing “hudson”

If HudsonNY is not installed, Illustrator will:
• Alert you
• Fall back to the default font

Restart Illustrator after installing fonts.

⸻

User Prompts

On run, the script prompts once for:
• Customer
• Vehicle

Vehicle Unit # is never prompted and always left blank for pen entry.

⸻

How to Use 1. Open the Illustrator document containing:
• One PROOF artboard
• One or more print artboards 2. Run:

FleetSpecSheet_V1.jsx

    3.	Enter:
    •	Customer
    •	Vehicle
    4.	Script will:
    •	Export artboards internally
    •	Create new SPEC SHEET artboard(s)
    •	Lay everything out automatically

⸻

Output
• New artboards added to the right of the document
• Named:
• SPEC SHEET (V1)
• SPEC SHEET (V1) — Pg 2, etc.
• One layer per page for cleanliness

⸻

Known Constraints (By Design)
• Artboard names are used verbatim for NAME / FILE columns
• Quantities must be encoded in artboard names
• No external file I/O (everything stays inside Illustrator)
• No auto-editing of source artboards

⸻

Version History

V1
• First stable Illustrator implementation
• Artboard-driven workflow
• Matches Python generator layout & logic
• Production-ready

⸻

Future Enhancements (Optional)
• Strip QTY\_\* from NAME column (display-only)
• Auto-detect proof by largest artboard if PROOF missing
• Sort manifest rows (by artboard order or name)
• Optional per-item notes column
• Expanded install sign-off options
