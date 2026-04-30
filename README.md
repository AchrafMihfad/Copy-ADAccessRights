# Copy-ADAccessRights

PowerShell script to copy NTFS permissions from one user to another.

## 🎯 Use case

When an employee leaves or changes position, a new user often needs the same access to network folders.

Instead of manually reassigning permissions folder by folder, this script automates the process.

## ⚙️ Features

- Copy NTFS permissions from one user to another
- Saves time and reduces human errors
- Works in Active Directory environments

## 🧪 Example

```powershell
.\Copy-ADAccessRights.ps1 -SourceUser "user1" -TargetUser "user2"
