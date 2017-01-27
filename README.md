# Description

Batch converts `.doc` to `.docx` using Microsoft Office 2011 for Mac.

# Install

- Microsoft Office 2011 for Mac
- copy `doc2docx.app` to `Applications`

# Usage

Either:

- GUI: drag a file or folder to the `doc2docx.app`
- CLI: `open <path_to>/doc2docx.app --args <folder>`

Alternatively, there's a lightweight `doc2docx.workflow`:

```bash
automator -i <path-to-folder> <path-to>/doc2docx.workflow
```

In addition, there are 2 workflows that uses Microsoft Word 2011 for Mac's action for conversion (the above uses Applescript), put in the folder `achive/`. But I found that the Microsoft Office 2011 provided automator component doesn't really convert doc to docx in the way that

1. file size does not change.
2. pandoc cannot read it.

Because of that, the achived scripts is of limited use.

Note: a `ppt2pptx.app` is newly added, but it assumes all files in the folder are presentation without checking it @todo

# Why not 2016?

- Microsoft Word 2016 is a regression in terms of both `.doc` handling and automator integration. YMMV if you use 2016 instead.

# Source

- [how do I use automator to batch convert doc to ... | Official Apple Support Communities](https://discussions.apple.com/thread/3050596?start=0&tstart=0)
- [The action "Convert Format of Word Documents" could not be - Microsoft Community](https://answers.microsoft.com/en-us/msoffice/forum/msoffice_word-mso_mac/the-action-convert-format-of-word-documents-could/d5b56a9a-d227-4f74-b3df-8f97377c6273)

# License

Since it was initially posted on a public forum, I assume it is in public domain. Consult the author in the Apple forum above.
