# SharePoint Site Setup Requirements

Voor de Presales Pipeline schema (`sharepoint_schema_definition.json`) moeten de volgende SharePoint site features en instellingen actief zijn.

---

## 1. Site Collection Features (Tenant Admin)

These require **SharePoint Online admin center** access.

### ✓ Required Features

| Feature | Scope | Reason |
|---------|-------|--------|
| **SharePoint Server Publishing Infrastructure** | Site Collection | Enables advanced content type management and versioning |
| **SharePoint Server Publishing** | Web (optional) | Needed if using publishing pages/web parts |

**How to enable:**
1. Go to Site Settings → Site Collection Features
2. Search for "Publishing Infrastructure"
3. Click **Activate**

---

## 2. Site-Level Features (Site Admin)

### ✓ Essential Features

| Feature | Reason |
|---------|--------|
| **Versioning** | Track document changes in DOCX library |
| **Content Types** | Assign metadata to lists/libraries |
| **Metadata Navigation** | Enable filtering by custom columns |
| **Item-level Permissions** | Restrict access to sensitive rows (optional) |

**How to enable:**
1. Go to Site Settings → Manage Site Features
2. Activate each feature

---

## 3. List & Library Configuration

### 3.1 Lists (`PresalesSDDocuments`, `PresalesSDTMapping`, `PresalesImages`, `PresalesSDT Mapping`, `PresalesPipelineRuns`, `PresalesPipelineIssues`)

#### ✓ Column Features Required

- **Text / Choice / Number / DateTime / URL** → Default support ✓
- **Indexed Columns** → Must create index for each (improves query performance)
- **Lookups** → Not used here, but if you want to link to other lists later

#### ✓ Settings per List

```
List Settings → Versioning Settings:
  □ Enable item versioning? → YES
  □ Keep drafts? → YES (allows soft deletes for audit trail)
  □ Require content approval? → NO (unless HITL review required)

List Settings → Advanced Settings:
  □ Allow management of content types? → YES
  □ Offline client availability? → YES
  □ Search results? → YES (for discovery)
```

#### Index Creation (Performance Tuning)

For each indexed column, create an index:
```
List Settings → Indexed columns → New Index
  Column: [SDName, RunId, FillType, Severity, etc.]
  Create lookup index? YES
```

**Why:** SharePoint view threshold is 5,000 items. Indexed columns bypass this limit.

### 3.2 Document Library (`PresalesPipelineArtifacts`)

#### ✓ Settings

```
Library Settings → Versioning Settings:
  □ Create a version each time you edit a file? → YES
  □ Keep versions? → 10 (retain last 10 versions)
  □ Require content approval? → NO (unless enforcing review gate)
  □ Document hold? → NO (unless legal compliance needed)

Library Settings → Advanced Settings:
  □ Opening Documents in the Browser? → Decide per format
    - DOCX → Open in browser OR local (your choice)
    - JSON/CSV/XLSX → Open in browser (SharePoint preview)

Library Settings → Folder Organization:
  □ Enable folder creation? → YES
  □ Folders: raw_json, mapped_json, xml, docx, dashboard, mapping_xlsx, images
```

#### Metadata Columns for Library

Apply columns from `sharepoint_schema_definition.json` `libraries[0].metadataColumns`:
- SDName (text, indexed)
- RunId (text, indexed)
- ArtifactType (choice, indexed)
- ImageId (text, optional)
- ImageFormat (choice, if artifact is image)
- ImageDimensions (text, if artifact is image)
- EmbeddedStatus (choice, if artifact is image)

```
Library Settings → Columns → Add from existing site columns
  (Or create new columns if not yet in site)
```

---

## 4. Microsoft Graph API Permissions (Tenant Admin)

For programmatic access via Python (Microsoft Graph SDK), register an **Azure AD app** with these permissions:

### ✓ Delegated Permissions (if user context)
- `Sites.ReadWrite.All` → Create/update lists & libraries
- `Files.ReadWrite.All` → Upload/download artifacts
- `User.Read` → Identify current user

### ✓ Application Permissions (if daemon/background process)
- `Sites.ReadWrite.All` → Unattended list operations
- `Files.ReadWrite.All` → Unattended file operations

**Setup:**
1. Azure Portal → App registrations → New registration
2. API Permissions → Add → Microsoft Graph
3. Select permissions above → Grant admin consent

---

## 5. Content Types Setup (Optional but Recommended)

Create site-level content types to standardize schema:

### ✓ Content Type: "Presales Image"
```
Name: Presales Image
Base: Item
Columns:
  - SDName (text)
  - RunId (text)
  - ImageId (text)
  - Format (choice)
  - Width, Height (number)
  - ImageSource (choice)
  - EmbeddedStatus (choice)
```

**How to create:**
1. Site Settings → Site Content Types
2. Create → Fill in fields
3. Assign to PresalesImages list

---

## 6. SharePoint Search & Indexing

### ✓ Enable Crawling

```
Site Settings → Search Settings:
  □ Allow this site to appear in search results? → YES
  □ Index? → YES (crawlers will index all public items)
```

**Impact:** Users can search for SDs by ProductCode, SDName, etc.

---

## 7. Scripting & Automation (If Using Microsoft Graph)

### ✓ Allow Custom Scripts (Advanced)
```
Site Settings → Site Collection App Catalog:
  (Only needed if deploying SharePoint add-ins; not required for Graph API)
```

**Note:** Most modern SharePoint automation uses Graph API directly (not SPFx).

---

## 8. Folder Structure in `PresalesPipelineArtifacts` Library

Manually create folders or automate via Python:

```
PresalesPipelineArtifacts/
  ├── raw_json/          → SD raw extraction JSON
  ├── mapped_json/       → Taxonomy-mapped JSON
  ├── xml/               → Structured XML output
  ├── docx/              → Generated presales DOCX files
  ├── dashboard/         → HTML & XLSX dashboards
  ├── mapping_xlsx/      → Per-SD SDT mapping exports
  └── images/
      ├── SD_Name_1/
      └── SD_Name_2/           (optional sub-organization)
```

---

## 9. SharePoint Online Limits to Know

| Limit | Value | Impact |
|-------|-------|--------|
| **List items per view** | 5,000 | Use filters + indexes; or batch by RunId |
| **File size (uploaded)** | 250 MB | No issue for DOCX/JSON/images |
| **Storage per site** | 1 TB (flexible) | Monitor DOCX + images accumulation |
| **List columns** | 500 | We use ~35; plenty of headroom |
| **Concurrent requests** | 100/min | Throttle Graph API calls if syncing 100s of items |

---

## 10. Checklist: Before First Data Sync

- [ ] Site collection created (e.g., `/sites/presales-pipeline`)
- [ ] Publishing Infrastructure feature activated
- [ ] Lists created: PresalesSDDocuments, PresalesSDTMapping, PresalesImages, PresalesPipelineRuns, PresalesPipelineIssues
- [ ] Library created: PresalesPipelineArtifacts
- [ ] Folders in library created (raw_json, mapped_json, xml, docx, dashboard, mapping_xlsx, images)
- [ ] Indexed columns created for each list (SDName, RunId, FillType, Severity, Format, EmbeddedStatus, etc.)
- [ ] Metadata columns added to library
- [ ] Versioning enabled on all lists & library
- [ ] Azure AD app registered with Graph API permissions
- [ ] Python environment has `msgraph-core` + `azure-identity` installed
- [ ] Tenant admin has granted consent to app permissions

---

## 11. Python Requirements for Graph Sync

```bash
pip install msgraph-core azure-identity openpyxl
```

### Authentication Methods

**Option A: User Context (Interactive)**
```python
from azure.identity import InteractiveBrowserCredential
credentials = InteractiveBrowserCredential(client_id="<APP_ID>")
```

**Option B: Service Principal (Background Jobs)**
```python
from azure.identity import ClientSecretCredential
credentials = ClientSecretCredential(
    tenant_id="<TENANT_ID>",
    client_id="<CLIENT_ID>",
    client_secret="<CLIENT_SECRET>"
)
```

---

## 12. Common Issues & Fixes

| Issue | Cause | Fix |
|-------|-------|-----|
| "List view threshold exceeded" | >5,000 items + no index | Create indexed column; filter view by RunId |
| "403 Unauthorized on list update" | Missing Graph permissions | Grant admin consent in Azure AD |
| "Cannot create nested folders" | Library doesn't allow folders | Enable folder creation in Library Settings |
| "Metadata column not showing" | Column not added to list/library | Add from Site Columns or create new |
| "Image embedding fails" | SharePoint doesn't support inline images | Store as URL metadata; reference from library |

---

## Next Steps

1. **Manual Setup** (if first time):
   - Create site collection
   - Activate features
   - Create lists/library
   - Run through checklist above

2. **Automated Setup** (recommended):
   - Use PnP PowerShell script or Python Graph SDK
   - Script can provision all lists/libraries + columns from JSON schema
   - See: `sharepoint_provision_schema.py` (to be created)

3. **Data Sync**:
   - Once setup complete, run sync script after each `run_all.py` execution
   - Upsert data into SharePoint lists

---

**Schema Reference:** [sharepoint_schema_definition.json](sharepoint_schema_definition.json)
