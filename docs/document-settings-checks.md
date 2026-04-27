# Document Settings Checks

The Document Settings checks show how a document is configured. These settings often influence how a document behaves when opened, edited, or saved, and can be critical for maintaining consistency, security, or specific workflows.

---

## Fields Checked under Document Settings

The following table details the specific document setting fields our plugin analyzes:

| Field Name             | Type       |    Access | Description |
| :------------------- | :-----------  | :----------- | :---------------------------------------------------------------------------------------------------------------------- |
| Compatibility Mode   | Bool / String | R | Indicates if the document is saved in an older compatibility mode, or details specific compatibility settings - `Compatibility Mode` in the title bar of the document. |  
| Auto Update Styles   | Bool         | R/W | Specifies if styles in the document are set to automatically update based on the attached template - 'Automatically update document styles' found in `Developer > Document Templates > Templates`     |          
| Anonymisation     | Bool         | R/W |Indicates whether personal information has been removed from the document properties to anonymize the file - `RemovePersonalInformation` or `RemoveDateAndTime` enabled in the developer properties. | 
| Document Protection  | Bool / String | R/W | Specifies if the document has protection enabled (e.g., read-only, password-protected, restricted editing), or details the type of protection. - `Restrict Editing` enabled in the Review tab. |

---
## Write Access
### Document Protection
When changing document protection via `Change attributes` the user will be provided with a prompt:

<img width="482" height="242" alt="image" src="https://github.com/user-attachments/assets/580ca38f-de6c-4514-9c0e-e3fa4687c296" />


Choose the protection type in the first dropdown, and a password in the second field. If the document is already protected, Total Commander will first remove protection then apply the new protection mode.

---
[Back to Documentation Overview](index.md)
