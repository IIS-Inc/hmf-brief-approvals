# HMF Brief Approval System
### Word-Based Approval Workflow with Quickbase Integration

---

## Overview

The HMF Brief Approval System is a Microsoft Word macro-enabled document (.docm) 
that provides a structured, role-based approval workflow for HMF Program Briefs. 
Approvals are managed through a custom ribbon interface embedded in the document 
and update Quickbase records in real time via the Quickbase REST API.

The workflow is linear and enforced by the system:

    Digital Review → Research Review → Executive Review

Each stage must be completed before the next becomes available.

---

## Repository Structure

    hmf-brief-approvals/
    ├── README.md
    ├── vba/
    │   ├── modUtilities.bas      # Core utilities, RID harvesting, document properties
    │   ├── modQuickbase.bas      # Quickbase REST API engine
    │   ├── modApprovals.bas      # Approval logic, workflow rules, UserForm controller
    │   ├── modRoles.bas          # Passcode authentication, role management
    │   ├── modRibbon.bas         # Custom ribbon callbacks
    │   └── modUpdater.bas        # GitHub-based auto-update utility (Admin only)
    ├── forms/
    │   └── frmApprovals.frm      # Approval UserForm
    ├── ribbon/
    │   └── customUI14.xml        # Custom ribbon XML definition
    └── template/
        └── HMF_Approval_Template.docm   # The production Word template

---

## Prerequisites

### Microsoft Word
- Microsoft Word for Office 365 (Version 16.0 or later)
- Macro-enabled documents (.docm) must be permitted in your environment

### Macro Security Settings
Each user must enable macros when prompted on document open, or the document 
must be placed in a Trusted Location.

### Trust Access to VBA Project Object Model
Required only for the Admin auto-update feature. Not needed for regular users.

To enable:
1. Open Word
2. File → Options → Trust Center → Trust Center Settings
3. Click Macro Settings in the left panel
4. Check "Trust access to the VBA project object model"
5. Click OK

NOTE: This is a per-machine setting. It does not travel with the document.
Only enable this on machines used by the Admin role.

### SharePoint / Network Considerations
- The document can be stored and opened from SharePoint
- Opening from SharePoint may trigger Protected View — users must click 
  Enable Editing then Enable Content to activate macros
- Adding the SharePoint site to Trusted Sites in Windows Internet Options 
  suppresses Protected View for that domain — recommended for all team members
- AutoSave is automatically enabled for SharePoint-hosted documents — 
  this is handled gracefully by the VBA code

---

## Quickbase Configuration

### Application
    Realm:   xxxxxxxxxxxxxxxxxxx.quickbase.com
    App ID:  xxxxxxxxxx

### Briefs Table
    Table ID: xxxxxxxxxx

    Field Label               Type                    Field ID
    ─────────────────────────────────────────────────────────
    Brief Approval Status     Text - Multiple Choice   8
    Digital Approval Status   Text - Multiple Choice   10
    Research Approval Status  Text - Multiple Choice   11
    Executive Approval Status Text - Multiple Choice   12
    Approval Log              Text                     13

### Roles Table
    Table ID: xxxxxxxxxx

    Field Label   Type                    Field ID
    ──────────────────────────────────────────────
    Email         Text                    6
    Role          Text - Multiple Choice  7
    Active        Checkbox                8

### API Authentication
The Quickbase user token is stored as a private constant in modQuickbase.bas.

NOTE: For production deployment, replace the development token with a 
production token and implement a secure token management strategy.
See the Security Considerations section below.

---

## Approval Workflow

### Status Values

#### Brief Approval Status (FID 8) — Master Status

    Value                               Stage
    ─────────────────────────────────────────
    Pending Digital review              1.0
    Digital: Editing                    1.1
    Digital: Rejected                   1.2
    Digital: Approved                   1.5
    Pending Research review             2.0
    Research: Editing                   2.1
    Research: Rejected                  2.2
    Research: Needs legal review        2.3
    Research: Approved                  2.5
    Pending Executive review            3.0
    Executive: Needs digital edits      3.1
    Executive: Needs Research/Legal edits 3.1
    Executive: Rejected                 3.2
    Executive: Approved                 3.5

#### Department Status Fields (FIDs 10, 11, 12)
Each department field stores the unprefixed status value (e.g. Approved, 
Editing, Rejected). The master Brief Approval Status (FID 8) stores 
the prefixed value (e.g. Digital: Approved).

### Workflow Gates
The system enforces the linear workflow programmatically:
- Digital buttons are active from stage 0 through 1.x
- Research buttons become active only after Digital reaches stage 1.5 (Approved)
- Executive buttons become active only after Research reaches stage 2.5 (Approved)

Attempting to act out of sequence displays an informational message and 
takes no action.

---

## HMF Record ID (HMF ID)

Each Brief document is linked to its corresponding Quickbase record via a 
numeric Record ID (RID). This is the Quickbase Record ID# field (FID 3) 
from the Briefs table.

### How It Works
1. When a document is opened for the first time, the system checks for a 
   stored RID in the document's Custom Properties
2. If no RID is found, the user is prompted to enter one
3. The RID is validated as a positive integer — non-numeric values are rejected
4. Once validated, the RID is saved to the document's Custom Properties and 
   persists for all future sessions
5. All Quickbase API calls use this RID to update the correct record

### Finding the RID
The Record ID can be found in Quickbase by:
- Opening the Brief record — the RID appears in the URL as rid=XXX
- Viewing the Record ID# field on the Brief form

### Clean Template
The production template must not contain a pre-stored RID. Before distributing 
the template, verify that all HMF Custom Properties have been cleared:

    File → Info → Properties → Advanced Properties → Custom tab

The following properties should all be absent or empty in a clean template:

    Property              Description
    ──────────────────────────────────────────────────────
    HMF_RID               The Quickbase Record ID
    HMF_BriefStatus       Cached Brief Approval Status
    HMF_DigitalStatus     Cached Digital status
    HMF_ResearchStatus    Cached Research status
    HMF_ExecutiveStatus   Cached Executive status
    HMF_UserRole          Cached user role for the session

To clear all properties programmatically, run from the VBA Immediate Window:

    CleanTemplate

---

## Passcode Authentication

Access to the approval ribbon is controlled by a passcode entered on every 
document open. The passcode determines which approval buttons are visible 
to the user.

### How It Works
1. Every time the document is opened, a passcode prompt appears
2. The user enters their assigned passcode
3. The system performs an exact case-sensitive match
4. The matched role determines which ribbon buttons are displayed
5. If no passcode is entered, or the passcode does not match, the document 
   opens in read-only mode with no approval buttons visible — no error is shown

### Roles and Ribbon Visibility

    Role       Digital Button   Research Button   Executive Button
    ──────────────────────────────────────────────────────────────
    Digital    Visible          Hidden            Hidden
    Research   Hidden           Visible           Hidden
    Executive  Hidden           Hidden            Visible
    Admin      Visible          Visible           Visible
    None/Wrong Hidden           Hidden            Hidden

### Passcode Management
- Passcodes are stored as private constants in modRoles.bas
- Passcodes are case-sensitive and should be treated as passwords
- For production deployment, use complex passcodes that are not dictionary words
- Passcodes should be distributed to team members securely — 
  never via email or chat in plaintext
- To change a passcode, update the constant in modRoles.bas, 
  push to GitHub, and run the Admin updater on the production template

IMPORTANT: Do not commit production passcodes to the GitHub repository.
Use placeholder values in the repository and maintain production 
values in a secure password manager.

---

## Admin Functions

### Auto-Update from GitHub
The modUpdater module allows an Admin to pull the latest VBA modules 
directly from the GitHub repository without manual file imports.

Requirements:
- Admin passcode must be entered on document open
- Trust access to the VBA project object model must be enabled on the machine

To run an update:

    UpdateFromGitHub

NOTE: modUpdater cannot update itself. If modUpdater.bas changes in the 
repository, it must be manually re-imported via the VBA Editor.

### Other Admin Utilities
Run these from the VBA Immediate Window (Alt+F11 → Ctrl+G):

    Command                    Description
    ──────────────────────────────────────────────────────────────────
    CleanTemplate              Removes all HMF Custom Properties
    ResetRole                  Clears cached role — prompts on next open
    SaveRIDToProperty 577      Manually sets the RID to a specific value
    TestConnection             Tests the Quickbase API connection
    TestPasscode               Tests the passcode flow without reopening
    UpdateFromGitHub           Pulls latest modules from GitHub (Admin only)

---

## Installing the Modules

### VBA Modules (.bas files)
1. Open the .docm in Word
2. Press Alt+F11 to open the VBA Editor
3. File → Import File
4. Navigate to the vba/ folder in your local repository clone
5. Select the .bas file and click Open
6. Repeat for each module
7. Rename modules if needed via the Properties panel (F4)

### UserForm (frmApprovals)
The UserForm cannot be imported the same way as a .bas file:
1. In the VBA Editor click Insert → UserForm
2. Rename it frmApprovals in the Properties panel (F4)
3. Add controls per the layout specification in forms/frmApprovals.frm
4. Double-click the form and paste the form code

### Custom Ribbon (customUI14.xml)
The ribbon XML must be added using the Office RibbonX Editor:

1. Download and install Office RibbonX Editor:
       https://github.com/fernandreu/office-ribbonx-editor/releases
2. Close the .docm in Word — the file must not be open
3. Open Office RibbonX Editor
4. File → Open → select the .docm
5. Right-click the file in the left panel
6. Select Insert Office 2010+ Custom UI Part
7. Paste the contents of ribbon/customUI14.xml into the editor
8. Click Validate — confirm no errors
9. Click Save then close RibbonX Editor
10. Reopen the .docm in Word

---

## Security Considerations

### Quickbase User Token
- The development token is stored in modQuickbase.bas as a private constant
- For production, consider one of the following approaches:
  - Runtime prompt — token entered once per session, never written to disk
  - Windows Credential Manager — token stored in the OS credential vault, 
    retrieved via Windows API calls from VBA
- Never commit a production token to a public repository
- Rotate tokens immediately if they are accidentally exposed

### Passcodes
- Passcodes provide role-based UI access control, not data security
- They do not encrypt the document or prevent access to the VBA code
- For sensitive deployments, consider VBA project password protection:
      VBA Editor → Tools → VBA Project Properties → Protection tab
- Production passcodes should never appear in the repository

### SharePoint
- Store the template in a SharePoint library with appropriate access controls
- Restrict edit access to the template file itself to Admin role only
- End users should work from copies, not the master template

---

## Development Workflow

### Branching Strategy
- main — production-ready code only
- dev — active development
- Feature branches for significant changes

### Recommended Commit Convention
    feat: add modUpdater - GitHub auto-update for Admin role
    fix: correct HTTP method in modQuickbase POST call
    docs: update README with SharePoint deployment notes
    refactor: simplify passcode logic in modRoles

### Testing Before Commit

    Module          Test Command                    Expected Result
    ──────────────────────────────────────────────────────────────────────
    modUtilities    InitializeDocument              RID prompt, status initialized
    modQuickbase    TestConnection                  HTTP 200 OK from Quickbase
    modRoles        TestPasscode                    Role assigned correctly
    modApprovals    LaunchApprovalForm "Digital"    UserForm opens
    modRibbon       Open doc with Admin passcode    All ribbon buttons visible

---

## Support and Administration

### Adding a New User
1. Distribute the passcode for their role securely
2. Ensure their machine has macros enabled for the SharePoint location

### Changing a Passcode
1. Update the constant in modRoles.bas
2. Commit and push to GitHub
3. Run UpdateFromGitHub on the production template with Admin passcode
4. Save the updated .docm back to SharePoint
5. Distribute the new passcode to affected users securely

### Rotating the Quickbase Token
1. Generate a new token in Quickbase:
       Profile → My Preferences → Manage User Tokens
2. Update QB_USER_TOKEN in modQuickbase.bas
3. Commit and push to GitHub
4. Run UpdateFromGitHub on the production template
5. Save the updated .docm back to SharePoint

---

## Version History

    Version   Date         Notes
    ────────────────────────────────────────
    1.0.0     2026-06-08   Initial development build

---

HMF Brief Approval System — developed for HMP
Repository: https://github.com/IIS-Inc/hmf-brief-approvals

