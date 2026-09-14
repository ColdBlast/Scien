# Scien
A spreadsheet built for science and engineering — not finance (a lightweight prototype)

Spreadsheets were designed for finance.  But scientists and engineers use them extensively, and it often feels like we are trying to shove a square peg into a round hole.
Wouldn't it be great if there were a spreadsheet package designed for science and engineering, from the ground up?  One that understands (and propagates) units and uncertainty?
One that generates scatter plots by default, with painless, point-by-point error whiskers?  One that understands regression and curve fitting at a higher level?

This isn't going to be that...but it's a fast-fashion shot at demonstrating that it is indeed possible, and that there's an audience for it!
It uses the framework of tabs within a conventional spreadsheet to emulate storing the multi-dimensional data required.  It can do some basic things, and I'd love to see it grow...
but the real goal is for an entity with much more muscle to do this right - from the ground up.
Right now it's aimed at the students I teach, an intro-level tool with intro-level capabilities; more of a teaching tool than a workhorse.  But try it and you'll see the promise!

I've had this idea rattling around in my head for a long time, but now I have the time to pursue it.  The last time I considered doing so, I was both inspired and defeated by
Matt McCutchen's public domain Measurements extension (https://mattmccutchen.net/measurements/), because it demonstrated possibility but also limited pursuit to a passion project.
The unit handling in Scien is inspired by Matt's approach, but Scien otherwise approaches things differently, and uses none of his code.  Gemini wrote most of Scien; I'm no coder.

Scien runs in Google Sheets as an Add-In, coded in Apps Script.  You can try it out by making a copy of this example Sheet: 
https://docs.google.com/spreadsheets/d/1OX9U-Q2s5-xgb5073HYfs_S_NYIMx0b1qOU3S8K355Y/copy
You type input into Scien naturally, with a value, uncertainty, and a unit, though you can use "+-" in place of ±; it is adaptive to spacing. There is some delay as it processes.
Scien currently supports only a limited set of common functions: +, -, *, /, ^, SQRT, AVERAGE, and STDEV.  These have not been extensively debugged: use with caution!

I am new at all of this, so will get things set up here only slowly.
