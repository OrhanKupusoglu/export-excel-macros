# Export MS Excel Macros

Since MS Excel workbooks are binary file formats, they are not suitable for revision control.

A [Bash script](./pre-commit/pre-commit) and a [Python script](./pre-commit/pre-commit.py) together export macros contained in workbooks.

These scripts are tested only on MS Windows.

&nbsp;

## Requirement

The Python tool [oletools](https://pypi.org/project/oletools/) is required to export the contained VBA forms & modules:

```
>pip install -U oletools
```

&nbsp;

## Git Hooks

The [Pro Git book](https://git-scm.com/book/en/v2) explains [Git Hooks](https://git-scm.com/book/en/v2/Customizing-Git-Git-Hooks) in great detail:

> Like many other Version Control Systems, Git has a way to fire off custom scripts when certain important actions occur.

&nbsp;

## Pre-Commit Hook

A **pre-commit** [hook](https://git-scm.com/book/en/v2/Customizing-Git-Git-Hooks) checks first for presence of unwanted worksheets. If such a worksheet is not present, then the pre-commit hook exports VBA modules to **src.vba** directory, so that all source files can be put under revision control.

The pre-commit hook script is taken from the blog post [How to use Git hooks to version-control your Excel VBA code](https://www.xltrail.com/blog/auto-export-vba-commit-hook).
The scripts are modified slightly.

The same company's open source [git-xl](https://github.com/xlwings/git-xl) extension for Git handles direct diff ops between workbook revisions, but it is not much of help inside an IDE.

These two files must first be put into the **.git/hooks** directory of the repository where workbooks are located.

&nbsp;

## Filter Worksheets

The shell script can give an *unwanted* worksheet name.
If it is empty filtering is ignored.

```
WS_UNWANTED='Confidential'
```

A worksheet with exactly this name or containing this name causes the commit to fail.
The Python script may check for exact matches, by default a partial match filters unwanted worksheets.

```
EXACT_MATCH = False
```

