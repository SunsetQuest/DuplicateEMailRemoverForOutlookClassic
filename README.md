# Duplicate Email Remover for Outlook (Classic)

## Overview
**Duplicate Email Remover for Outlook Classic** is a Windows Forms application that connects to Microsoft Office Outlook (Classic) and scans the user's mailboxes and PST files for duplicate emails. The application provides multiple options for handling duplicates, including deletion, moving, copying, logging, and exporting. The software was created in November 2023 and updated in December 2024.

## Key Features
- **Folder Selection**: Users can choose specific folders or mailboxes to scan for duplicates.
- **Customizable Folder Order**: Adjust the processing order of folders to control where duplicates are removed.
- **Flexible Actions**: Options include deleting, moving, copying, logging, or opening duplicates in Outlook.
- **Advanced Matching Options**: Match duplicates based on fields like date, sender, subject, body, and attachments.
- **Backup Capability**: Save duplicates to a specified Windows file folder for backup.
- **Compatibility**: Designed for Outlook Classic and does not support the "New Outlook" version.

## System Requirements
- **Microsoft Office Outlook Classic** installed.
- **.NET 8 LTS** runtime.
- A Windows environment.

## Installation

### Method 1: Manual Build
1. Clone the repository:
   ```bash
   git clone https://github.com/SunsetQuest/DuplicateEMailRemoverForOutlookClassic.git
   ```
2. Open the solution file `DuplicateEMailRemoverForOutlookClassic.sln` in Visual Studio.
3. Build the solution and run the application.

### Method 2: Prebuilt Executable
1. Visit [Releases](https://github.com/SunsetQuest/DuplicateEMailRemoverForOutlookClassic/releases).
2. Download the latest version.

## Usage

### Step-by-Step Guide
1. **Select Folders**
   - Start with the "Select Folders" tab to choose which mailboxes and folders to scan.
   - Right-click on a folder to include all its subfolders.
   - System skips folders like `Calendar`, `Contacts`, and other non-email folders.

2. **Arrange Folder Order**
   - In the "Folder Order" tab, drag and drop folders to set their processing order.
   - Folders processed earlier retain their emails; duplicates in later folders are removed.

3. **Define Actions**
   - Select an action for duplicates in the "Action to Take" tab:
     - Move to a specified folder.
     - Copy to a specified folder.
     - Open each email in Outlook.
     - Delete duplicates.
     - Log duplicates to a file only.
   - Optionally save each duplicate to a Windows file folder.
   - Choose matching criteria from fields like `DateTime Sent`, `Sender Address`, `Subject`, etc.

4. **Accept Terms**
   - Read and accept the license agreement and warnings on the "Accept" tab.

5. **Start Processing**
   - Click "Start" on the "Go" tab to begin scanning and handling duplicates.
   - View the status of folders and duplicates in real-time.
   - A log file of actions taken will open automatically upon completion.

### Notes
- **Do not open or close Outlook while the tool is running.**
- Outlook does not need to be open to use this tool.

## ChatGPT/AI Usage
As of 12/22/2024, all source code in this project was created entirely by Ryan Scott White without the assistance of AI tools. This project originated in November 2023, during a period when Ryan relied solely on traditional IT resources. However, this README.md file was crafted with the help of ChatGPT-4 on 12/21/2024.

## Note Regarding Co-Pilot
Microsoft currently does not recognize me as an "open source contributor," so I have not had access to their $10/month CoPilot service. I did have the opportunity to use CoPilot for a year in 2021/2022, during its pre-release phase when it was offered for free to all users. I often wonder whether my 20+ original projects, publicly available on GitHub, contributed to training the AI. Given this body of work, I feel I should be acknowledged as an open source contributor.

## License
This project is licensed under the MIT License. See the `LICENSE` file for details.

## Disclaimer
The software is provided "as is," without warranty of any kind. Users should back up important data before using the tool. The author is not liable for any data loss or damage.

---

Enjoy a cleaner mailbox with **Duplicate Email Remover for Outlook Classic**!
