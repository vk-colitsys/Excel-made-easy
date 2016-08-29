<cfscript>
  /**
  * @description Populate workbook sheets with data
  * @output Yes
  */
  public struct function populateWorkbook(required struct doc_struct){
    doc = ARGUMENTS.doc_struct;

    // Loop through each HTML table
    for(t=1; t LTE ArrayLen(doc.tables); t=t+1) {
      doc.current_row = 1;
      table = doc.tables[t];
      table_attributes = table.XmlAttributes;

      // Define Name of Spreadsheet Tab
      table_name = ( isDefined('table_attributes.name') ? table_attributes.name : "Sheet_#t#" );
      table_name = REReplace(table_name, "[^a-zA-Z0-9\']", "_", "ALL");

			/*
				Define where the data starts
				(very first cell that is NOT a row heading or column heading)
			*/
      doc.data_origin = getDataOrigin(table);

      if(t EQ 1) { doc.first_sheet_name = table_name; }

      // Create current sheet (with name)
      SpreadsheetCreateSheet(doc.workbook, table_name);

      // Remove default first sheet if exists
      try { SpreadsheetRemoveSheet(doc.workbook, "Sheet1"); } catch(any e) {}

      // Make modifications to current sheet
      SpreadsheetSetActiveSheet(doc.workbook, table_name);

      /* ========== HEAD ==================================================== */
      if (isDefined('table.thead.tr')) {
        doc = populateWorksheetSection(doc, table.thead.tr);
      }//end:HEAD

      /* ========== BODY ==================================================== */
      if (isDefined('table.tbody.tr')) {
        doc = populateWorksheetSection(doc, table.tbody.tr);
      }//end:BODY

      /* ========== FOOTER ================================================== */
      doc.footer_start_row = doc.current_row;
      if (isDefined('table.tfoot.tr')) {
        doc = populateWorksheetSection(doc, table.tfoot.tr);
      }//end:FOOT
      doc.footer_rows = doc.current_row-doc.footer_start_row;
      doc.last_data_row = doc.current_row-1;

      /* ========== TWEAKS ================================================== */
      doc = tweakWorksheetDisplay(doc);

      /* ========== FORMATTING ============================================== */
      doc = themeWorksheet(doc, table);
      
      /* ========== AUTOMATIC COLUMN WIDTHS ================================= */
      doc = setColumnWidths(doc, table);
    }//end:tables

    return doc;
  }//end:populateWorkbook()

  /**
  * @description Populate the three possible sections of a table
  * @output Yes
  */
  public struct function populateWorksheetSection(required struct doc_struct, rows){
    //doc = StructCopy(doc_struct);
    doc = ARGUMENTS.doc_struct;
    for( r=1; r LTE ArrayLen(ARGUMENTS.rows); r+=1 ){
      row = ARGUMENTS.rows[r];
      doc.current_column = 1;

      if (isDefined('row.th')) { doc = populateRowSegment(doc, row.th); }
      if (isDefined('row.td')) { doc = populateRowSegment(doc, row.td); }

      doc.current_row += 1;
    }//end:rows loop

    return doc;
  }//end:populateWorksheetSection()

  /**
  * @description Populate a section of a row
  * @output Yes
  */
  public struct function populateRowSegment(required struct doc_struct, cells){
    //doc = StructCopy(doc_struct);
    doc = ARGUMENTS.doc_struct;

    for( c=1;  c LTE ArrayLen(ARGUMENTS.cells); c+=1 ) {
      this_cell = ARGUMENTS.cells[c];
      SpreadsheetSetCellValue(
        doc.workbook,
        Trim(this_cell.XmlText),
        doc.current_row,
        doc.current_column
      );
      doc.current_column += 1;
    }//end:cells Loop

    return doc;
  }//end:populateRowSegment()

</cfscript>