---
name: sharepoint
description: "SharePoint file operations and Office document intelligence via Microsoft Graph API. Use when: (1) listing, reading, uploading, or searching files in SharePoint, (2) extracting text from Word/Excel/PowerPoint/PDF documents stored in SharePoint, (3) cleaning up meeting notes into professional documents, (4) summarizing documents from business or technical perspectives, (5) creating Excel trackers or PowerPoint presentations from notes. Requires certificate-based auth with Sites.Selected permission."
---

# SharePoint Skill

Interact with SharePoint document libraries via Microsoft Graph API using certificate-based authentication.

## Operations

### File operations

```bash
# List files in library root
node scripts/sharepoint.mjs list

# List files in a subfolder
node scripts/sharepoint.mjs list --path "Meeting Notes/2026"

# Read file content (extracts text from Office formats)
node scripts/sharepoint.mjs read --path "Meeting Notes/2026-02-23.docx"

# Upload a file
node scripts/sharepoint.mjs upload --local ./report.docx --remote "Reports/Q1-2026.docx"

# Search for files
node scripts/sharepoint.mjs search --query "quarterly review"

# Create folder
node scripts/sharepoint.mjs mkdir --path "Meeting Notes/2026"

# Delete (requires --confirm flag)
node scripts/sharepoint.mjs delete --path "Drafts/old-file.txt" --confirm
```

### Reading Office documents

The `read` command extracts text content from supported formats:
- `.docx` → full text extraction via mammoth
- `.xlsx` → sheet names + cell data via exceljs
- `.pptx` → slide text extraction via pptxgenjs
- `.pdf` → text extraction via pdf-parse
- `.txt` / `.md` → raw content

Output is plain text suitable for AI processing (summarization, reformatting, etc.).

## Configuration

Environment variables (or `.env` file):

```
SP_TENANT_ID=<azure-tenant-id>
SP_CLIENT_ID=<app-client-id>
SP_CERT_PATH=<path-to-pfx-or-key>
SP_CERT_PASSWORD=<pfx-password-if-any>
SP_SITE_ID=<sharepoint-site-id>
SP_DRIVE_ID=<optional-specific-drive-id>
```

## Security controls

- Path traversal prevention: all paths are canonicalized, `../` is rejected.
- File size limit: configurable max download/upload size (default 50MB).
- Operation allowlist: delete is disabled unless `--confirm` flag is passed.
- No tokens or secrets in logs.
- Certificate auth only (no client secrets).

## Setup guide

For complete setup from scratch (Entra app, certificate, Sites.Selected, Key Vault):
See `references/setup-guide.md`
