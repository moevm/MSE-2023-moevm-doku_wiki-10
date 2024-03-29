# XLSX2DW Plugin for DokuWiki
The XLSX2DW plugin for DokuWiki makes it easy to import XLS, XLSX or ODS tables to a page.

The plugin keeps the styles, merged cells and colors of your original table.

## Plugin installation
1.1. Download from DokuWiki Extension Manager ([Plugin page link](https://www.dokuwiki.org/plugin:xlsx2dw)).

1.2. Clone repository (via SSH or HTTPS) or download from DokuWiki Extension Manager:
- SSH: `git clone git@github.com:moevm/MSE-2023-moevm-doku_wiki-10.git`
- HTTPS: `git clone https://github.com/moevm/MSE-2023-moevm-doku_wiki-10.git`

Move all plugin files to `dokuwiki-installation-directory/lib/plugins/xlsx2dw/`
- TIP: `dokuwiki-installation-directory` is folder of your local DokuWiki.

After all files have been transferred, the `/xlsx2dw` directory should look like the image below.

![Alt text](./screenshots/plugin_directory.png?raw=true "/xlsx2dw folder")

## Usage
1. Create new page (or edit the page).

![Alt text](./screenshots/creating_page_section.png?raw=true "Create page")

2. Click "Import table" button in the toolbar.

![Alt text](./screenshots/using_button.jpg?raw=true "Import of tables")

3. Choose a table file. You can select example table from `/_test/test-tables/` folder.

![Alt text](./screenshots/selecting_tables.png?raw=true "Select a table")

4. Selected table is converted to DokuWiki syntax.

![Alt text](./screenshots/table_in_dokusyntax.png?raw=true "DokuWiki syntax")

5. The result is shown in the screenshot below.

![Alt text](./screenshots/preview_table.png?raw=true "Result")
