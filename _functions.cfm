<cfscript>
	include "_theming.cfm";
	include "_formatting.cfm";
	include "_populating.cfm";
	include "_supporting.cfm";

  /**
  * @description This is the primary function for generating an Excel Document
  * @output Yes
  * @hint Return true if successfully generated Excel document
  */
  public boolean function table2excel(required struct args){
    fn_args = ARGUMENTS.args;

    // DEFAULT SWITCHES/ARGUMENTS
    param fn_args.tables = "";
    param fn_args.theme = "default";
    param fn_args.file_location = "";
    param fn_args.auto_freeze_headers = false;
    param fn_args.freeze_at = "";
    param fn_args.column_widths = "auto";

    // Boolean Switches for Header Theming
    param fn_args.theme_data_cells = false;
    param fn_args.theme_column_headings = true;
    param fn_args.theme_row_headings = true;
    param fn_args.theme_footer = true;

    // Use one or the other (if both, use custom)
    param fn_args.theme = "default";
    param fn_args.custom_theme = ""; // TODO: Option to use JSON

    if (fn_args.tables EQ '') { return false; }
    if (fn_args.file_location EQ "") { return false; }

		// Determine Excel Document Format
    isXmlFormat = ( getFiletypeFromPath(fn_args.file_location) EQ "Excel2007" ? true : false );

    // New Spreadsheet Object
    workbook = SpreadsheetNew("Sheet1", isXmlFormat);

    /*
      This step is crucial because it will make any HTML character entities
      valid XML entities even though they don't all import into Excel.
    */
    var args_tables = makeValidXML(args.tables);

    /* Fix any lone ampersands that may cause failed XML parsing */
    args_tables = fix_broken_entities(args_tables);

		/*
			Initialize the doc struct

			This struct is passed throughout the process and contains useful
			information needed for each step.
		*/
		doc = {
      workbook = workbook,
      current_row = 1,
      tables = XMLParse(args_tables).tables.table,
      first_sheet_name = "Sheet1",
      config = fn_args
    };

		// Populate Workbook with HTML Table Data
    doc = populateWorkbook(doc);

		// First Sheet active when opened
    SpreadsheetSetActiveSheet(doc.workbook, doc.first_sheet_name);

    try {
      // Save Spreadsheet
      SpreadsheetWrite(doc.workbook, args.file_location, 'yes');
      return true;
    } catch(any e) { return false; }
  }//end:table2excel()
</cfscript>
