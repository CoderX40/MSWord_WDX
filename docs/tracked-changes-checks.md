# Tracked Changes Checks

The Tracked Changes checks are specifically designed to evaluate the presence and nature of revisions within documents. This is incredibly valuable for documents that undergo collaborative editing, providing a historical record of modifications.

---

## Fields Checked under Tracked Changes

The following table details the specific fields the plugin analyses regarding tracked changes:

| Field Name                       | Type       | Description                                                                                                                                                                                                                                                                                                 |
| :------------------------------- | :--------- | :---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Track Changes Activated/Deactivated** | Bool       | Indicates whether the "Track Changes" feature was activated or deactivate when the document was last saved. This doesn't necessarily mean changes are present, only that the feature was active.                                                                                                            |
| **Tracked Changes Present** | Bool       | Indicates whether the document contains any actual tracked changes (insertions, deletions, moves, or formatting changes) that have not yet been accepted or rejected.                                                                                                                                          |
| **Tracked Changes Authors** | String     | A list or string of unique authors who have contributed tracked changes to the document. (e.g., "John Doe; Jane Smith").                                                                                                                                                                                           |
| **Total Revisions** | Int        | The cumulative count of all individual revisions (insertions, deletions, moves, formatting changes) present in the document.                                                                                                                                                                                    |
| **Insertions** | Int        | The total number of text or object insertions recorded as tracked changes.                                                                                                                                                                                                                                  |
| **Deletions** | Int        | The total number of text or object deletions recorded as tracked changes.                                                                                                                                                                                                                                   |
| **Moves** | Int        | The total number of text or object movements recorded as tracked changes (where content was cut from one place and pasted into another, and tracked as a move).                                                                                                                                                 |
| **Formatting Changes** | Int        | The total number of formatting modifications (e.g., bolding, font changes, paragraph spacing) recorded as tracked changes.           |

*Note: The number of Moves and Formatting changes may vary slightly compared to the amount show in the Revisions Pane in Word, due to how Word interprets these changes to show the user in the file.*

---

[Back to Documentation Overview](index.md)
