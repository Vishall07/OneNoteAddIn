# OneNote Bookmarks Add-in
<img width="1362" height="780" alt="image" src="https://github.com/user-attachments/assets/6d482b46-186b-4dad-8639-4f4ddfc65766" />

A custom OneNote add-in that adds a **Bookmarks** menu to the OneNote ribbon. This tool allows users to save, organize, and quickly access OneNote objects (notebooks, sections, pages, paragraphs) through a hierarchical bookmarks menu.

## Features

### Core Features

- **Ribbon Integration:**  
  Adds a "BOOKMARKS" dropdown menu to the OneNote ribbon to access saved bookmarks.

- **Bookmarking:**  
  Save current notebook, section group, section, page, or paragraph as a bookmark with automatic labels.

- **Persistent Storage:**  
  Bookmarks and folder organization persist after closing OneNote or restarting the PC.

- **Folders & Organization:**  
  Create unlimited nested folders with full create, rename, delete, and drag-and-drop support.  
  Folders sort alphabetically and save their open/close state.

- **Bookmark Management:**  
  Rename bookmarks in-place, delete with confirmation, reorder, multi-select with Ctrl/Shift for batch operations, and drag bookmarks between folders.

- **Menu Behavior:**  
  Resizable menu that stays open after adding bookmarks until the user clicks outside; right-click options for word wrap and column settings.  
  Menu closes on activating a bookmark, clicking elsewhere in OneNote, or pressing Escape.

- **Keyboard Navigation:**  
  Arrow keys to navigate, F2 to rename focused item, Enter to open focused item.

### Optional Features

- Display icons for notebooks, section groups, sections, pages, and paragraphs with matching colors and symbols.

- Right-click option to open all bookmarks in a folder in separate OneNote windows.

- Export bookmarks list to plain `.txt` file.

- Drag and drop URLs from browsers like Firefox to add as bookmarks.

- Free-text notes column with in-place editing and word wrap toggle.

- Path column showing full OneNote hierarchy path (read-only), with word wrap toggle.

## Installation

- This add-in is a standalone internal tool for the full OneNote desktop app included with Microsoft 365 (not Store/UWP app).

- Install by loading the add-in DLL into OneNote through Visual Studio or manually registering the COM add-in if applicable.

## Usage

- Use the **BOOKMARKS** dropdown on the OneNote ribbon to add bookmarks for the current notebook, section group, section, page, or paragraph.

- Organize bookmarks in folders using drag-and-drop or right-click context menus.

- Double-click or press Enter on a bookmark to open it in the active OneNote window.

- Use keyboard shortcuts for quick navigation and editing of bookmarks.

## License

This tool is for internal use and not published in any public store.
