# Word as basic Latex editor/previewer

> There is a dichtonomy: Those who use Latex and those who use Word. Let us get over it!

This Word-.doc-file, which has VBA-Macros (only tested on MS Windows), tries to address the following problem:
* Some people want to write papers with Latex, but supervisors and co-authors do not want to work with Latex; they work solely in Word
* You want to profit from Word: Word has good collaboration features, very tidy and flexible track change and commenting system: You can comment on comments, concurrent editing etc. You can also compare two documents, with a reasonable, editable(!) merged file. The spelling and grammar checking is good. You can fold headings.
* Converting a latex-generated .pdf to word is easy: Just open the .pdf in Word. The problem is the back-conversion to Latex, and the usual program `pandoc` cannot handle track-changes etc. done by co-authors well. So, you usually have to manually insert all the tiny changes in the revised .doc back to your .tex file. `Overleaf` is great, but the track changes and commenting is fair at the moment. 

Hence, we try to stay totally with Latex-Code in Word without any conversions or rich-text substitutions.

The current VBA-Macros are basic, there are still many issues, and the code needs extensions, but the VBA-code works under these limitations.

The VBA-code (see `LatexEditor.doc` or `LatexEditor.pdf`, with the helper file `preamble.tex` for the Latex-preamble of snippets) let you:
* Preview displayed content inside word (ehh, latex): Activated environment at the moment: `equation`, `align`, `table`, `figure`. Currently, you can preview and delete all generated pictures at once only.
* Format some common Latex-code to gray or even white, and format the sections into Word-headings.
* Run Latex from Word

For more documentation, see `LatexEditor.doc` or `LatexEditor.pdf`, or look into the VBA-code (which you have to do anyway to adjust a path and to unlock the shell-code, which is usually detected by anti-virus programs as dangerous code, usually while saving the .doc)
