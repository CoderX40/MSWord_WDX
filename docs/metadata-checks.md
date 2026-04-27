# Metadata Checks

The Metadata checks focus on properties embedded within documents that provide information about the document's creation, authorship, and general context. These fields are often automatically populated by the creating application or document template used, but can also be manually edited.

Metadata can generally be found in Word documents under `File > Info > Properties > Advanced Properties`

---

## Fields Checked under Metadata

The following table details the specific metadata fields our plugin analyzes:

| Field Name           | Type         | Access | Description 
| :------------------- | :----------- | :----- | :-----------------------------------------------------------------------------------------------------
| **Title** | String      | R/W | The main title or name given to the document. Quickly identify document purpose; ensure consistent naming conventions.|
| **Subject** | String       | R/W| A brief description or summary of the document's content.|
| **Author** | String       | R/W| The primary creator or author of the document.  |
| **Manager** | String      | R/W | The person responsible for the document's content or lifecycle. (Common in some enterprise systems).|
| **Company** | String      | R/W | The organisation associated with the document's creation or ownership.                 |                
| **Keywords** | String      | R/W | Tags or terms that describe the document's content for search and categorization purposes.  |            
| **Comments** | String      | R/W | General comments or notes related to the document as a whole (distinct from in-document comments).   |
| **Template** | String      | R | The path and filename of the template used to create the document.                     |               
| **Hyperlink Base** | String       | R/W| The base path for relative hyperlinks within the document.                   |                      
| **Created** | DateTime     | R/W| The date and time when the document was originally created.                         |              
| **Modified** | DateTime    | R/W | The date and time when the document was last saved or modified.                   |        
| **Printed Dates** | DateTime    | R/W | The date and time when the document was last printed.                                |                
| **Last Modified By** | String / Int | R/W| The user or system that last saved or modified the document. Can be a name or a system ID.  |           
| **Revision Number** | String / Int | R/W| A number or identifier tracking the document's revision count. Incremented with each save.   |       
| **Total Editing Time** | Int         | R/W | The cumulative time (in minutes or seconds) spent editing the document.          |               
| **Pages** | Int          | R | The total number of pages in the document.                    |                                       
| **Words** | Int          | R | The total count of words within the document's main content.   |                                     
| **Characters** | Int      | R | The total count of characters (including spaces) within the document's main content.|               
| **Paragraphs** | Int       | R | The total number of paragraphs in the document.                                    |                    
| **Lines** | Int         | R | The total number of lines in the document. (Note: May vary based on word wrap/display settings).     |

---

[Back to Documentation Overview](index.md)
