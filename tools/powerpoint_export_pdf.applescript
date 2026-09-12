-- Export a .pptx to PDF using the locally installed Microsoft PowerPoint.
--
-- Used as the ground-truth oracle for visual regression testing: PowerPoint is the
-- reference implementation, so its own output is the only fully authoritative answer
-- to "what should this slide look like?".
--
--   osascript tools/powerpoint_export_pdf.applescript <input.pptx> <output.pdf>
--
-- Then rasterise per page with pypdfium2.
--
-- Constraints discovered on PowerPoint 16.x / macOS:
--   * PowerPoint is sandboxed. Both paths must be somewhere it can reach -- the user's
--     home tree works, /tmp and /private/tmp fail with error -9074.
--   * Paths must be absolute; the presentation is matched by its exact full name so a
--     concurrently open deck is never exported by mistake.
--   * `save as PNG` exists in the dictionary but silently produces nothing. PDF is the
--     only export that works unattended; per-slide PNG needs a VBA macro host.
--   * The first run triggers a macOS automation permission prompt.
--   * A damaged file raises an app-modal repair dialog ("Herstellen" / "Repair").  While it
--     is up PowerPoint stops servicing AppleEvents, so this script cannot dismiss it -- it is
--     blocked inside `open`.  The timeout below turns a 120-second hang into a prompt
--     failure; clearing the dialog afterwards needs a separate process (see the recovery
--     helper in pptx-agent's tests/oracle.py).  A timeout here therefore means "PowerPoint
--     would not open this file", which is exactly the verdict the oracle exists to give.

on run argv
    if (count of argv) is not 2 then
        error "Usage: osascript powerpoint_export_pdf.applescript <input-pptx> <output-pdf>"
    end if

    set inputPosixPath to item 1 of argv
    set inPath to POSIX file inputPosixPath
    set outPath to POSIX file (item 2 of argv)

    set openedPresentation to missing value
    -- Fail fast instead of waiting out the 120-second default: a stall here means a modal
    -- dialog, not slow work, and no amount of extra waiting clears one.
    with timeout of 45 seconds
    tell application "Microsoft PowerPoint"
        activate
        try
            open inPath
            -- Match by full path rather than trusting the active window.
            set presentationPaths to (get full name of every presentation)
            repeat with presentationIndex from 1 to count of presentationPaths
                if (item presentationIndex of presentationPaths as text) is inputPosixPath then
                    set openedPresentation to presentation presentationIndex
                    exit repeat
                end if
            end repeat
            if openedPresentation is missing value then
                error "PowerPoint opened the input but no presentation matched: " & inputPosixPath
            end if
            save openedPresentation in outPath as save as PDF
            close openedPresentation saving no
        on error errorMessage number errorNumber
            if openedPresentation is not missing value then
                try
                    close openedPresentation saving no
                end try
            end if
            error errorMessage number errorNumber
        end try
    end tell
    end timeout
end run
