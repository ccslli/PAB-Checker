# ---- PAB Checker: a 0-100% bar instead of the file names (as in Design Team Tools) --
# Appended to PyInstaller's splash-screen script by pab_checker.spec
# (ProgressSplash).  The .exe's unpacker sets status_text to the name of each
# file as it extracts it - @TOTAL@ files in this build - which drives 0-90%.
# The program then sends "NN%|label" (pab_checker_gui.py: splash_progress) for the rest,
# reaching 100% as the main window opens.  The text always says "Loading..."
# (the labels are ignored - only the percentage is used).

# Wrapped in catch: if a PyInstaller version names things differently, the
# splash simply shows without the bar instead of failing.
if {[catch {
    trace remove variable status_text write [list canvas_text_update .root.canvas vartext]

    set pb(x0) 22
    set pb(x1) 458
    set pb(y0) 213
    set pb(y1) 217
    set pb(total) @TOTAL@
    set pb(count) 0
    set pb(pct) 0

    .root.canvas create rectangle $pb(x0) $pb(y0) $pb(x1) $pb(y1) -fill #2A3A42 -outline "" -tags pb_trough
    .root.canvas create rectangle $pb(x0) $pb(y0) $pb(x0) $pb(y1) -fill #3CC9CB -outline "" -tags pb_fill
    .root.canvas create text $pb(x1) 236 -anchor se -fill #9AA7B4 -font myFont -text "0%" -tags pb_pct
    .root.canvas coords vartext $pb(x0) 236
    .root.canvas itemconfigure vartext -text "Loading..."

    proc pb_show {pct label} {
        global pb
        if {$pct < $pb(pct)} { set pct $pb(pct) }
        if {$pct > 100} { set pct 100 }
        set pb(pct) $pct
        set x [expr {$pb(x0) + ($pb(x1) - $pb(x0)) * $pct / 100.0}]
        .root.canvas coords pb_fill $pb(x0) $pb(y0) $x $pb(y1)
        .root.canvas itemconfigure pb_pct -text "$pct%"
        .root.canvas itemconfigure vartext -text "Loading..."
    }

    proc pb_update {args} {
        global pb status_text
        if {[regexp {^(\d+)%\|(.*)$} $status_text -> pct label]} {
            pb_show $pct $label
        } else {
            incr pb(count)
            pb_show [expr {min(90, int(90.0 * $pb(count) / max(1, $pb(total))))}] ""
        }
    }

    trace add variable status_text write pb_update
} pb_error]} {
    set pb_failed $pb_error
}
