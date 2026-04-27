# Tracked Changes Checks

The Tracked Changes checks are specifically designed to evaluate the presence and nature of revisions within documents. This is incredibly valuable for documents that undergo collaborative editing, providing a historical record of modifications.

---

## Fields Checked under Tracked Changes

The following table details the specific fields the plugin analyses regarding tracked changes:

| Field Name                       | Type      | Access | Description                                                                                                                                                                                                                                                                                                 |
| :------------------------------- | :--------- |:------ | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Track Changes Mode** | Bool      | R/W | Indicates whether the "Track Changes" feature was enabled or disabled when the document was last saved. This doesn't necessarily mean changes are present, only that the feature was enabled.                                                                                                            |
| **Contains Tracked Changes** | Bool      | R | Indicates whether the document contains any actual tracked changes (insertions, deletions, moves, or formatting changes) that have not yet been accepted or rejected.                                                                                                                                          |
| **Tracked Changes Authors** | String    | R/W | A list or string of unique authors who have contributed tracked changes to the document. (e.g., "John Doe; Jane Doe").                                                                                                                                                                                           |
| **Total Revisions** | Int       | R | The cumulative count of all individual revisions (insertions, deletions, moves, formatting changes) present in the document.                                                                                                                                                                                    |
| **Total Insertions** | Int      | R  | The total number of text or object insertions recorded as tracked changes.                                                                                                                                                                                                                                  |
| **Total Deletions** | Int       | R | The total number of text or object deletions recorded as tracked changes.                                                                                                                                                                                                                                   |
| **Total Moves** | Int       | R | The total number of text or object movements recorded as tracked changes (where content was cut from one place and pasted into another, and tracked as a move).                                                                                                                                                 |
| **Total Formatting Changes** | Int       | R | The total number of formatting modifications (e.g., bolding, font changes, paragraph spacing) recorded as tracked changes.           |

*Note: The number of Moves and Formatting changes may vary slightly compared to the amount show in the Revisions Pane in Word, due to how Word interprets these changes to show the user in the file.*

---
## Write Access
### Tracked Changes Authors
When changing tracked changes authors via `Change attributes` the user will be provided with a prompt:

<img width="522" height="272" alt="image" src="https://github.com/user-attachments/assets/db575da7-3106-4649-a3ad-c87e8e1f57ef" />


Enter the author to replace in the first field, and the new user author in the second field. If the first field is left blank, then all authors across files will be updated.

---
[Back to Documentation Overview](index.md)
