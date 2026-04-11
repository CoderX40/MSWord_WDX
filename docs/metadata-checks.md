# Metadata Checks

The Metadata checks focus on properties embedded within documents that provide information about the document's creation, authorship, and general context. These fields are often automatically populated by the creating application or document template used, but can also be manually edited.

Metadata can generally be found in Word documents under `File > Info > Properties > Advanced Properties`

---

## Fields Checked under Metadata

The following table details the specific metadata fields our plugin analyzes:

| Field Name           | Type         | Description 
| :------------------- | :----------- | :-----------------------------------------------------------------------------------------------------
| **Title** | String       | The main title or name given to the document.| Quickly identify document purpose; ensure consistent naming conventions.|
| **Subject** | String       | A brief description or summary of the document's content.|
| **Author** | String       | The primary creator or author of the document.  |
| **Manager** | String       | The person responsible for the document's content or lifecycle. (Common in some enterprise systems).|
| **Company** | String       | The organisation associated with the document's creation or ownership.                 |                
| **Keywords** | String       | Tags or terms that describe the document's content for search and categorization purposes.  |            
| **Comments** | String       | General comments or notes related to the document as a whole (distinct from in-document comments).   |
| **Template** | String       | The path and filename of the template used to create the document.                     |               
| **Hyperlink Base** | String       | The base path for relative hyperlinks within the document.                   |                      
| **Created** | DateTime     | The date and time when the document was originally created.                         |              
| **Modified** | DateTime     | The date and time when the document was last saved or modified.                   |        
| **Printed Dates** | DateTime     | The date and time when the document was last printed.                                |                
| **Last Modified By** | String / Int | The user or system that last saved or modified the document. Can be a name or a system ID.  |           
| **Revision Number** | String / Int | A number or identifier tracking the document's revision count. Incremented with each save.   |       
| **Total Editing Time** | Int          | The cumulative time (in minutes or seconds) spent editing the document.          |               
| **Pages** | Int          | The total number of pages in the document.                    |                                       
| **Words** | Int          | The total count of words within the document's main content.   |                                     
| **Characters** | Int          | The total count of characters (including spaces) within the document's main content.|               
| **Paragraphs** | Int          | The total number of paragraphs in the document.                                    |                    
| **Lines** | Int          | The total number of lines in the document. (Note: May vary based on word wrap/display settings).     |

---

[Back to Documentation Overview](index.md)
