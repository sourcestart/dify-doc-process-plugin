# dify-doc-process-plugin

DOC processing tool plugin for Dify.

## Version

- Current version: `0.0.1`
- Capability scope: only `DOC -> DOCX`

## Feature

- Convert uploaded `.doc` files to `.docx`
- Return converted `.docx` directly as Dify file output (`blob message`)

## Tool

- Tool name: `doc_to_docx`
- Input:
  - `input_file` (`file`, required): must be a `.doc` file
- Output:
  - Converted `.docx` file
  - A JSON status message for workflow debugging

## Dependencies

- Python dependency:
  - `dify_plugin>=0.4.0,<0.7.0`
- External runtime dependency:
  - LibreOffice CLI (`soffice` or `libreoffice`) must be installed and available in `PATH`

## Error Handling

The tool returns explicit errors for:

- Missing uploaded file
- Uploaded file is not `.doc`
- Uploaded file is empty
- LibreOffice dependency missing
- Conversion command failure / timeout
- Output file missing or empty after conversion

## Temporary File Cleanup

The tool processes files in `tempfile.TemporaryDirectory()`. Temporary input/output files are automatically deleted after invocation.

## Usage Notes and Limits

- This version only supports `.doc -> .docx`.
- `.docx`, `.pdf`, and other formats are intentionally out of scope for `0.0.1`.
- Conversion quality depends on LibreOffice compatibility with the source `.doc`.
