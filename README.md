# Motion Document Organizer

This repository contains a small utility script to automatically sort
files from an inbox folder and generate markdown indexes compatible
with [Obsidian](https://obsidian.md/).

Run `organize_documents.py` to scan `Inbox-tris-docs` and move files
into the `Documents triés/` hierarchy. After organising, the script
creates a markdown file in each category to help navigation.

```
python3 organize_documents.py
```

The script only uses Python standard libraries and works on macOS.

