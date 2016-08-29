<cfscript>
  /**
  * @description Set Column Widths
  * @output Yes
  */
  public struct function setColumnWidths(required struct doc_struct, table){
    doc = ARGUMENTS.doc_struct;

    // Determine column_width array
    LOCAL.widths = getColumnWidthArray(table);

    // Set automatic
    if (
      doc.config.column_widths EQ "auto" OR
      doc.config.column_widths EQ ""
    ) {
      for (col=1; col LTE ArrayLen(local.widths); col+=1){
        adjusted_width = Round(1.5 * local.widths[col]);
        
        // Prevent invalid column widths
        if (adjusted_width LT 1) { adjusted_width = 1; }
        if (adjusted_width GT 255) { adjusted_width = 255; }
        
        SpreadsheetSetColumnWidth(doc.workbook,col,adjusted_width);
      }
    }
    
    // Set explicit
    if (
      doc.config.column_widths NEQ "auto" AND
      ListLen("#doc.config.column_widths#") GT 0
    ) {
      // Use Explicit Column Width Array
      explicit_widths = ListToArray(doc.config.column_widths, ',', true);
      
      // Fill in missing values
      for (i=1; i LTE ArrayLen(LOCAL.widths); i+=1) {
        if (NOT ArrayIsDefined(explicit_widths, i)) {
          explicit_widths[i] = LOCAL.widths[i];
        } else if ( NOT isNumeric(explicit_widths[i]) ) {
          explicit_widths[i] = LOCAL.widths[i];
        }
      }
      
      for (col=1; col LTE ArrayLen(explicit_widths); col+=1){
        adjusted_width = Round(1.25 * explicit_widths[col]);
        
        // Prevent invalid column widths
        if (adjusted_width LT 1) { adjusted_width = 1; }
        if (adjusted_width GT 255) { adjusted_width = 255; }
        
        SpreadsheetSetColumnWidth(doc.workbook, col, adjusted_width);
      }
    }
    
    return doc;
  }//end:setColumnWidths()


  /**
  * @description Get Array of Column Widths
  * @output Yes
  */
  public array function getColumnWidthArray(table){
    var out = [];
    
    // ========== CHECK THEAD ==================================================
    if ( isDefined("ARGUMENTS.table.thead.tr") ) {
      rows = ARGUMENTS.table.thead.tr;
      for(row=1; row LTE ArrayLen(rows); row+=1) {
        this_row = rows[row];
        for(col=1; col LTE ArrayLen(this_row.th); col+=1) {
          this_cell = this_row.th[col];
          cell_length = Len(this_cell.XmlText);
          if (NOT ArrayIsDefined(out, col)) { out[col] = cell_length; }
          if ( cell_length GT out[col] ) { out[col] = cell_length; }
        }//end:cols
      }//end:rows
    }//end:thead
    
    // ========== CHECK TBODY ==================================================
    if ( isDefined("ARGUMENTS.table.tbody.tr") ) {
      rows = ARGUMENTS.table.tbody.tr;
      for(row=1; row LTE ArrayLen(rows); row+=1) {
        this_row = rows[row];
        current_col = 1;
        
        if ( isDefined("this_row.th") ) {
          for(col=1; col LTE ArrayLen(this_row.th); col+=1) {
            this_cell = this_row.th[col];
            cell_length = Len(this_cell.XmlText);
            
            if (NOT ArrayIsDefined(out, current_col)) {
              out[current_col] = cell_length;
            }
            if ( cell_length GT out[current_col] ) {
              out[current_col] = cell_length;
            }
            current_col += 1;
          }
        }//end:th

        if ( isDefined("this_row.td") ) {
          for(col=1; col LTE ArrayLen(this_row.td); col+=1) {
            this_cell = this_row.td[col];
            cell_length = Len(this_cell.XmlText);
            
            if (NOT ArrayIsDefined(out, current_col)) {
              out[current_col] = cell_length;
            }
            if ( cell_length GT out[current_col] ) {
              out[current_col] = cell_length;
            }
            current_col += 1;
          }
        }//end:td

      }//end:rows
    }//end:tbody

    // ========== CHECK TFOOT ==================================================
    if ( isDefined("ARGUMENTS.table.tfoot.tr") ) {
      rows = ARGUMENTS.table.tfoot.tr;
      for(row=1; row LTE ArrayLen(rows); row+=1) {
        this_row = rows[row];
        for(col=1; col LTE ArrayLen(this_row.th); col+=1) {
          this_cell = this_row.th[col];
          cell_length = Len(this_cell.XmlText);
          if (NOT ArrayIsDefined(out, col)) { out[col] = cell_length; }
          if ( cell_length GT out[col] ) { out[col] = cell_length; }
        }//end:cols
      }//end:rows
    }//end:tfoot

    return out;
  }//end:getColumnWidthArray()

  /**
  * @description Use this to set formatting for an entire column (faster than individual cells)
  * @output Yes
  */
  public struct function setExplicitColumnFormatting(required struct doc_struct, table){
    doc = ARGUMENTS.doc_struct;
    if ( isDefined('ARGUMENTS.table.thead.tr') ) {
      head_row = ARGUMENTS.table.thead.tr[1];
      for(col=1; col LTE ArrayLen(head_row.th); col=col+1){
        cell = head_row.th[col];
        if ( isDefined('cell.XmlAttributes.style') ) {
          parsed_format = getFormatFromStyle("#cell.XmlAttributes.style#");
          hard_format = { color = "blue" };
          if ( NOT StructIsEmpty(parsed_format)) {
            SpreadsheetFormatColumns(doc.workbook, parsed_format, "#col#");
          }
        }
      }
    }
    return doc;
  }//end:setExplicitColumnFormatting()

  /**
  * @description Use this to set formatting for an entire row
  * @output Yes
  */
  public struct function setExplicitRowFormatting(required struct doc_struct, table){
    doc = ARGUMENTS.doc_struct;
    if ( isDefined("ARGUMENTS.table.thead.tr") ) {
      heading_height = ArrayLen(ARGUMENTS.table.thead.tr);
    } else {
      heading_height = 0;
    }
    if ( isDefined('ARGUMENTS.table.tbody.tr') ) {
      rows = ARGUMENTS.table.tbody.tr;
      for(r=1; r LTE ArrayLen(rows); r=r+1) {
        row = rows[r];
        if ( isDefined('row.XmlAttributes.style') ) {
          parsed_format = getFormatFromStyle(row.XmlAttributes.style);
          SpreadsheetFormatRow(doc.workbook, parsed_format, "#r+heading_height#");
        }
      }
    }
    return doc;
  }//end:setExplicitRowFormatting()


  /**
  * @description Get cell-specific formatting from style attribute
  * @output Yes
  * XXX: This is only working if the css comment is at the start of the html style attribute
  * TODO: Fix above so it works anywhere in html style attribute
  */
  public struct function getFormatFromStyle(string style){
    out = {};
    start = REFind("/\*", style)+2;
    end = REFind("\*/", style)-start;
    if (start GT 0 and end GTE start) {
      special_mid = Trim(Mid(style, start, end));
      pairs = ListToArray(special_mid, ";");
      for(i=1; i LTE ArrayLen(pairs); i=i+1) {
        pair = ListToArray(pairs[i], ":");
        key = Trim(pair[1]);
        value = Trim(pair[2]);
        StructInsert(out, key, value, true);
      }
    }
    return out;
  }//end:getFormatFromStyle()

</cfscript>
