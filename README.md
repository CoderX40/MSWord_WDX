# 📄 Microsoft Word Content Plugin (.wdx) for Total Commander

A C++ `.wdx` plugin that enables Total Commander to display detailed metadata and tracked change statistics from Microsoft Word `.docx` files - without needing Microsoft Word installed.

## ✨ Features

This plugin extends Total Commander by allowing you to:

* **View document metadata**: Title, subject, author, company, revision number, and more.
* **Inspect tracked changes**: See total insertions, deletions, moves, and formatting changes.
* **Detect document protection**: Check if protection or anonymisation is enabled.
* **Count comments** and **hidden text** in documents.
* Works directly with `.docx` files using native parsing - no Office dependency.

---

## 🔧 Supported Fields

| Category              | Field Name                        | Type       |
|-----------------------|-----------------------------------|------------|
| **Metadata**          | Title, Subject, Author, Manager, Company, Keywords, Comments, Template, Hyperlink Base | String |
|                       | Created, Modified, Printed Dates | DateTime   |
|                       | Last Modified By, Revision Number | String / Int |
|                       | Total Editing Time, Pages, Words, Characters, Paragraphs, Lines | Int |
| **Document Settings** | Compatibility Mode, Auto Update Styles, Anonymised Files, Document Protection | Bool / String |
| **Document Checks**          | Number of Comments                | Int        |
|                       | Hidden Text                   | Bool       |
| **Tracked Changes**   | Track Changes Activated/Deactivated, Tracked Changes Present | Bool |
|                       | Tracked Changes Authors           | String     |
|                       | Total Revisions, Insertions, Deletions, Moves, Formatting Changes | Int |


*Note: Fields appear in Total Commander’s "Custom Columns" dialog when configuring a content plugin view.*

---

## 💻 Installation

1. **Download**: Get the latest `.wdx` plugin file from the [Releases](https://github.com/CoderX40/MSWord_WDX/releases).
2. **Install in Total Commander**:
   - Open Total Commander → `Configuration > Options > Plugins`.
   - Under **Content Plugins (.WDX)**, click **Configure...**, then **Add...**.
   - Select the plugin `.wdx` file. Confirm installation.
3. **Create a custom column view**:
   - `Configuration > Options > Custom columns`.
   - Add a new column set, click **Add column**, and select fields from the plugin.

---

## 📂 Usage

Once installed, navigate to a folder with `.docx` files:

- Open the custom column view you created.
- The document properties will be displayed in the file list as new columns.


---

## 🛠️ Building from Source

### ✅ Prerequisites

- Windows
- Visual Studio (C++ Desktop Development workload)
- Dependencies:
  - [tinyxml2](https://github.com/leethomason/tinyxml2) (for XML parsing)
  - [miniz](https://github.com/richgel999/miniz) (for `.docx` unzip support)

### 🔨 Build Steps

```bash
git clone https://github.com/CoderX40/MSWord_WDX.git
cd MSWord_WDX
```
Rename the output `.dll` file with the extension `.wdx` or `.wdx64`.

## ⚠️ Notes & Limitations

* **No Microsoft Word required** – The plugin extracts data directly from `.docx` files.
* **Only `.docx` files** – This plugin does **not** support `.doc` (legacy binary) files.
* **Corrupted documents** – Some malformed `.docx` files may fail silently or return partial data.
* **Performance on large files** – Parsing large documents with lots of tracked changes or comments may cause a slight delay. Or viewing folders with several documents.
* **Field availability varies** – Some metadata fields (e.g. revision number, printed date) may not be present if not set in the document. Pages populates from a value directly within the `app.xml` file which may not show the correct pagecount for certain documents.
* **Tested on Total Commander 10+**, on Windows 10 and 11. Older versions may still work but are untested.

## 🤝 Contributing

Contributions are welcome! Whether you're fixing bugs, adding features, improving performance, or extending compatibility, feel free to jump in.

Bug reports and suggestions via GitHub Issues are also welcome.

## 📄 License

This project is licensed under the **MIT License**.

You are free to use, modify, distribute, and sublicense this plugin, as long as the original license and copyright notice are included.

See the [LICENSE](LICENSE) file for full details.
