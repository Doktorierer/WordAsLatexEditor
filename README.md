# Word as basic Latex editor/previewer

Use MS Word as a LaTeX editor, that is, you can rreview LaTeX environments, and send the .doc to co-authors and supervisors for revision, in case these people work usually with MS Word. The revised .doc can edited further in MS Word, or copy to your favourite LaTeX editor (Ctrl-A, Ctrl-C).


## Files

* `LatexEditor.doc`: Example LaTeX document in Word)
* `LatexEditor.pdf`: PdfLaTeX output of `LatexEditor.doc` (by clicking on "Run LaTeX")
* `template.tex`: The LaTeX premable for the preview of the snippets


## Problem addressed

The VBA-Macros (as in `LatexEditor.doc`) try to address the following problem:

* Some people want to write papers with Latex, but supervisors and co-authors do not want to work with Latex; they work solely in Word
* You want to profit from Word: Word has good collaboration features, very tidy and flexible track change and commenting system: You can comment on comments, concurrent editing etc. You can also compare two documents, with a reasonable, editable(!) merged file. The spelling and grammar checking is good. You can fold headings.
* Converting a latex-generated .pdf to word is easy: Just open the .pdf in Word. The problem is the back-conversion to Latex, and the usual program `pandoc` cannot handle track-changes etc. done by co-authors well. So, you usually have to manually insert all the tiny changes in the revised .doc back to your .tex file. `Overleaf` is great, but the track changes and commenting is fair at the moment. 

Hence, we try to stay totally with LaTeX-Code in Word without any conversions or rich-text substitutions.

## Features

The VBA-code (as in`LatexEditor.doc` with the helper file `preamble.tex` for the Latex-preamble of snippets) let you (from buttons on a new `LaTeX` tab):
* Preview LaTeX environments: Activated at the moment: `equation`, `align`, `table`, `figure`, etc. Currently, you can preview and delete all generated pictures at once only.
* Format some common LaTeX markup to gray or white, format the sections into Word headings, and hide LaTEX markup
* Run pdfLaTeX from Word and display the resulting .pdf

For more documentation, `LatexEditor.doc` (or `LatexEditor.pdf`)

## Prerequisites

* LaTeX installation (with latemk command)
* LaTeX packages: [TODO]

## Troubeshooting


* Tested on MS Windows 10 and Word 2016.
* If commented in the VBA-code, uncomment the lines with shell code. Such lines may be flagged by anti-virus software as suspicious code, especially during saving; see instruction in `LatexEditor.doc` for the lines
* If the code does not run and gives errors at creating objects, please ensure that the required References in the VBA code are ticked:
  * VBA Scripting Runtime
  * VBA Regex

## Caveats

The current VBA-Macros are basic, and the code needs extensions

