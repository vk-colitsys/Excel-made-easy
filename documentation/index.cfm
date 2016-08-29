<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
  <head>
    <title>magicExcel Documentation</title>
    <link rel="stylesheet" type="text/css" href="documentation.css" />
  </head>
  <body>
    <h1>magicExcel Documentation</h1>
    
    <table class="doctable" cellpadding="5px" cellspacing="0px" width="80%">
      <thead>
        <tr>
          <th>tag</th>
          <th>description</th>
        </tr>
      </thead>
      <tbody>
        <tr>
          <td><a href="build_from_table.html" target="_blank">build_from_table</a></td>
          <td>Generate an Excel doc by coding HTML tables. This will require html tables coded with certain conventions in place.</td>
        </tr>
        <tr>
          <td><a href="build_from_query.html" target="_blank">build_from_query</a></td>
          <td>Generate an Excel doc from a ColdFusion query result. Point it to a query result and presto!</td>
        </tr>
        <tr>
          <td><a href="build_from_json.html" target="_blank">build_from_json</a></td>
          <td>Generate an Excel doc from a JSON query format.</td>
        </tr>
        <tr>
          <td><a href="excel_color_formats.xls">Color Chart Reference</a></td>
          <td>
            <p>I've found that the fgcolor and color are the two colors you would want to use to style a spreadsheet.</p>
            <p>Interestingly, 'fgcolor' applies to the cell background while 'color' will apply to the text.</p>
            <p>There is a 'bgcolor' formatting option, but I have yet to figure out what it does.</p>
          </td>
        </tr>
      </tbody>
    </table>
    <br />
    
    <h2>Changelog</h2>
    <dl>
      <dt>Version 0.4.1</dt>
      <dd>New attribute <code>theme_data_cells</code> for <code>build_from_table</code></dd>
      <dd>Bug Fix
        <dl>
          <dd>
            (Related to new attribute of <code>build_from_table</code> tag.)<br />
            Data cell theming on very large spreadsheets has the potential to cause a memory leak and crash the CFML interpreter.<br />
            As such, data cells will NOT be themed by default and will need to be enabled explicitly using the <code>theme_data_cells</code> attribute for <code>build_from_table</code>.
          </dd>
        </dl>
      </dd>

      <dt>Version 0.4.0</dt>
      <dd>Bug Fixes
        <dl>
          <dd>Fixed missing entity error pertaining to lone ampersands in the table cell contents.</dd>
        </dl>
      </dd>

      <dt>Version 0.3.2</dt>
      <dd>bug fixes
        <dl>
          <dd>Coldfusion Sets a column width constraint of 255. Automatic and explicit column widths will not exceed this number.</dd>
          <dd>Because of the possibility of very wide columns, the headers will default to be left aligned.</dd>
        </dl>
      </dd>
      
      <dt>Version 0.3.1</dt>
      <dd>build_from_table improvements
        <dl>
          <dd>Automatic Column Widths</dd>
          <dd>Support for explicit column widths</dd>
          <dd>Fixed logic for applying styles to data rows</dd>
        </dl>
      </dd>
        
      <dt>Version 0.3.0</dt>
      <dd>Added build_from_query</dd>
      <dd>Added build_from_json</dd>
      <dd>Documentation begun</dd>
        
      <dt>Version 0.2.0</dt>
      <dd>RIAforge Import</dd>
        
      <dt>Version 0.1.0</dt>
      <dd>Initial Development &amp; Testing</dd>
    </dl>
    <br />
    
    
    <h2>To-Do/Wish List</h2>
    <dl>
      <dt>build_from_table</dt>
        <dd>
          <ul>
            <li>colspan/rowspan support</li>
            <li class="done">Explicit Column Widths</li>
            <li class="done">Auto Column Widths</li>
          </ul>
        </dd>
      <dt>build_from_csv</dt>
        <dd>
          <ul>
            <li>Make it</li>
            <li>Auto Column Widths</li>
          </ul>
        </dd>
      <dt>build_from_query</dt>
        <dd>
          <ul>
            <li>Auto Column Widths</li>
            <li class="done">Make it</li>
          </ul>
        </dd>
      <dt>build_from_json</dt>
        <dd>
          <ul>
            <li>ThisTag.generatedContent support</li>
            <li>Multi-sheet Support</li>
            <li class="done">Make it</li>
          </ul>
        </dd>
      <dt>build_from_textile</dt>
        <dd>
          <ul>
            <li>Make it</li>
          </ul>
        </dd>
    </dl>
  </body>
</html>
