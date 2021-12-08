# excel-regex-tool
A GUI for previewing your Regex pattern before applying it to data in Excel.

## How to Use

Enter your pattern, and result mask if required, then press Preview.

The preview will be populated with results from a small sample of the selection, including whether or not a cell matches your pattern. 

If you are happy with the preview, press Run to apply the pattern to the entire selection. The Results window will list how many cells passed the pattern, failed the pattern, or were blank, non-text, or errors.

If you are happy with the results, press Save to commit the data to the sheet.

Optionally, the Highlight checkbox will highlight pattern matches in green, and non-matches in cross-hatched pink.

## References
These are Early-bound for now, and need to be added via Tools > References

* Microsoft VBScript Regular Expressions 5.5 (vbscript.dll)
* Microsoft Windows Common Controls 6.0 (SP6) (mscomctl.ocx)

## To-do
* Gracefully handle malformed patterns
* Gracefully handle malformed result masks
* Apply results to a different column
* Store/recall previous patterns
* Implement unit testing