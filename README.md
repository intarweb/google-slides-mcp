# google-slides-mcp

Google Slides over MCP using an existing Google OAuth `authorized_user` refresh token (the file
google-workspace-mcp writes). The Slides API accepts the `drive` scope, so no new consent is needed;
the Google Cloud project behind the OAuth client must have the Slides API enabled.

Tools: `get_presentation`, `batch_update`, `replace_text`, `convert_office` (.pptx → native Slides),
`export_pptx` (native Slides → .pptx in Drive). The Slides API can't edit Office files directly:
convert first, then edit the copy.

    GOOGLE_SLIDES_CREDENTIALS=/path/to/authorized_user.json uvx --from git+https://github.com/intarweb/google-slides-mcp google-slides-mcp
