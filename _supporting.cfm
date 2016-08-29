<cfscript>
  /**
  * @description Determine excel type from file path
  * @output Yes
  */
  public string function getFiletypeFromPath(required string path){
    tmp = ListToArray(ARGUMENTS.path, ".");
    extension = tmp[ArrayLen(tmp)];

    switch(extension){
      case "xls": return "Excel2003"; break;
      case "xlsx" : return "Excel2007"; break;
      default: return "Excel2003"; break;
    }//end:switch:
  }//end:getFiletypeFromPath()


  /**
  * @description Inject XHTML entities into xml string
  * @output No
  *   This will make the character entities valid,
  *   but they will NOT import into Excel.
  */
  public any function makeValidXML(string the_xml){
    out = '<?xml version="1.0" encoding="ISO-8859-1"?>
      <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
      <tables>#the_xml#</tables>
    ';
    return out;
  }//end:injectXHTMLentities()

  /**
  *	@description Fix broken entities
  * @output No
  */
  public string function fix_broken_entities(required string the_string) {
		//var regex = '&(?=[^\b]*\b[^&\;]+\;?)';
		var regex = '&(?!\b[^&\;]+\;)';
		return ReReplace(ARGUMENTS.the_string, regex, "&amp;", "all");
  }//end:fix_broken_entities()

  /**
  * @description Generate Struct with extremely useful table information
  * @output No
  */
  public struct function getDataOrigin(xml_table){
    // Set Defaults
    row_header_depth = 0;
    col_header_height = 0;

    /* ========== Determine Row Header Depth ================================ */
    // Determine the number of Row Headers
    if (isDefined('ARGUMENTS.xml_table.tbody.tr')){
      rows = ARGUMENTS.xml_table.tbody.tr;
      for(i=1; i LTE ArrayLen(rows); i=i+1) {
        this_row = rows[i];
        if (isDefined('this_row.th')) {
          row_header_depth = ( ArrayLen(this_row.th) GT row_header_depth ? ArrayLen(this_row.th) : row_header_depth );
        }
      }
    }

    // Do we have a thead element?
    if ( isDefined("ARGUMENTS.xml_table.thead.tr") ) {
      col_header_height = ArrayLen(xml_table.thead.tr);
    }

    return { row = (col_header_height+1), column = (row_header_depth+1) };
  }//end:getDataOrigin()


  /**
  * @description Worksheet Display Tweeks
  * @output Yes
  */
  public struct function tweakWorksheetDisplay(struct doc_struct){
    doc = ARGUMENTS.doc_struct;

    if (doc.config.auto_freeze_headers) {
			// Freeze at Data Origin if automatically detecting freeze coordinates
			LOCAL.freeze_coord = [doc.data_origin.row, doc.data_origin.column];
    } else {
			// Use default (1,1)
			LOCAL.freeze_coord = ListToArray(doc.config.freeze_at);
		}


		if ( ArrayLen(LOCAL.freeze_coord) NEQ 2 ) { LOCAL.freeze_coord = [1,1]; }

		/*
			Neither the row nor column freeze coordinate can be less than 1

			Excel gives a "data maybe lost" error if we don't test this.
			This is not an issue when opened in OpenOffice

			A Freeze Pane should only be added if the coordinate is greater than 1,1.
		*/
    if (LOCAL.freeze_coord[1] GT 1 OR LOCAL.freeze_coord[2] GT 1) {

			SpreadsheetAddFreezePane(
				doc.workbook,
				LOCAL.freeze_coord[2]-1,
				LOCAL.freeze_coord[1]-1
			);
		}

    return doc;
  }//end:tweakWorksheetDisplay()

</cfscript>