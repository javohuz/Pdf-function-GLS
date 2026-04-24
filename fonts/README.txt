Place a Japanese font file here for the most reliable PDF output.

Recommended file names:
- NotoSansJP-Regular.ttf
- NotoSerifJP-Regular.otf
- NotoSansCJKjp-Regular.otf

You can also point the app to a font file with:
PDF_JP_FONT_PATH=/absolute/path/to/font.ttf

If no font file is provided, the app falls back to ReportLab's built-in
Japanese CID font. That fallback often works, but bundling a real font file
is better for consistent rendering across environments and PDF viewers.
