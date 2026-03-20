# skewt-vb6
Program to plot a Skew-T diagram and data, written in Visual Basic 6

By Cody Kirkpatrick

## Development machines

Developed using Visual Studio 6 on:
- a Compaq Presario 5170 (year 1998; Pentium II 350 MHz) running
  Windows 98 SE
- a Dell Dimension 2400 (year 2003; Celeron 2.4 GHz) running 
  Windows 98 SE, way over-spec for that era I know

## History & motivation

Overall development began March 15, 2026.  I had written a
QuickBasic program to plot skew-T data previously.  Since I've
wanted to learn VB for a long time now, this project plus spring
break seemed like the right time to dive in.

Google Gemini has aided me in this project.

Why keep this level of detail in versioning?  To keep track of
each of the individual features I learned about, added, had to
troubleshoot, and ultimately got working.  For next time!

## Versions

Major version 1, minor version 1

1.1.3
- Version number now displayed in title bar

- Image save name now matches input file, was previously a 
generic file name

- Cleaned up comments in source code, with more to be done

- This is the only revision of 1.1.x in which the source code
is available publicly

1.1.2
- Cleaned up button names & cleaned up some status messages, with
more to be done

1.1.1 - March 19, 2026
- Auto-polling (fixed at 10 sec) is enabled once a file is
plotted, but can be disabled 

Major version 1, minor version 0

1.0.9
- Cursor readout update: added a MouseMove event to reset the
hover readout when the cursor leaves the skew-T
- Audio: added annoying "ding" sounds when a plot successfully
finishes rendering
- This is the only revision of 1.0.x in which the source code
was preserved

1.0.8
- File type expansion: uses InStr to read two different types of
Intermet file types (TS vs. HS & HL).  Switches which columns to
read
- Data cutoff: Stop plotting at 16 km where the plot top is
reached, eliminating unnecessary file I/O

1.0.7
- Statistics: display surface and last-line-read values into a 
readout panel ("label") on the left

1.0.6
- Windows file integration: uses CommonDialog to browse and 
select data files instead of relying a single hardcoded
file name
- Button to replot the same file in case it is being updated live

1.0.5
- Coordinate readout: mouse-over the skew-T and see the T and
height of that pixel displayed on screen

1.0.4
- Added ability to save image by clicking a button

1.0.3
- Plot button to clear data out of the plot
- Plot base skew-T on startup

1.0.2
- Read in RH from InterMet files, subroutine to convert to dewpoint
- Added printing of mixing ratio lines to base Skew-T

1.0.1
- Plot basic skewed temperature lines, and create a button to
read in T(z) from a fixed file name


(end of document)
