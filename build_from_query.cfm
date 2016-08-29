<cfscript>
	/*
		AUTHOR: Ryan Johnson (rhino.citguy@gmail.com)
		LICENSE: GPLv3 (http://www.gnu.org/licenses/gpl.html)
		VERSION:
		DESCRIPTION:
			Custom Tag to transform a ColdFusion Query into a Microsoft Excel document.
	*/
  ThisTag.name = "BUILD_FROM_QUERY";

  param ATTRIBUTES.query = "";
  param ATTRIBUTES.file = "";
  
  param ATTRIBUTES.theme = "";
  param ATTRIBUTES.custom_theme = "";

  // variable name to set results in CALLER scope
  param ATTRIBUTES.variable = "build_from_query_result";
  
  // Various Checks prior to executing the code
  if (ATTRIBUTES.file EQ "") { throw(message="FILE is required for #ThisTag.name#"); }
  if (ATTRIBUTES.query EQ "") { throw(message="QUERY is required for #ThisTag.name#"); }

  // This tag will execute code at the end of its closing tag
  if (ThisTag.hasEndTag) {
    if (ThisTag.ExecutionMode EQ "start") {
    } else {
      engage();
    }
  } else {
    throw(message="#ThisTag.name# must have a closing tag. (You may self-close #ThisTag.name# if you wish.)");
  }

</cfscript>

<cffunction name="engage"
  access="public" returntype="any" output="Yes"
  description="I execute the logic for this tag."
  hint="Make it so, number one!"
>
  <cfscript>
    /*
      The following lines aren't necessary but are here as a placeholder for
      possible future development.
    */
    LOCAL.content = ThisTag.GeneratedContent; // Capture Content
    ThisTag.GeneratedContent = ""; // Prevent browser display

    include "_supporting.cfm";
    include "_theming.cfm";

    // Is the generated file in the new Excel XML format?
    isXmlFormat = ( getFiletypeFromPath(ATTRIBUTES.file) EQ "Excel2007" ? true : false );
    
    // New Spreadsheet Object
    LOCAL.workbook = SpreadsheetNew("Sheet1", isXmlFormat);
    
    // Serialize/Deserialize will turn the query into a struct
    LOCAL.json_content = SerializeJSON( Evaluate("CALLER.#ATTRIBUTES.query#") );
    LOCAL.struct_content = DeserializeJSON(LOCAL.json_content);

    // Parse through JSON to populate workbook
    LOCAL.workbook = populateFromJSON(LOCAL.workbook, LOCAL.json_content);
    
    // Used to determine row at which to stop formatting
    total_rows = ArrayLen(LOCAL.struct_content.data) + 1;
    
    /* ========== THEME THE WORKBOOK ======================================== */
    if (ATTRIBUTES.theme NEQ "" OR ATTRIBUTES.custom_theme NEQ "") {
      if (isStruct(ATTRIBUTES.custom_theme)) {
        useTheme = ATTRIBUTES.custom_theme;
      } else {
        useTheme = getFormatForTheme(ATTRIBUTES.theme);
      }
      // Format Header Row
      SpreadsheetFormatRow(LOCAL.workbook, useTheme.definition.heading, "1");
      
      // Format Data Rows
      SpreadsheetFormatRows(LOCAL.workbook, useTheme.definition.data, "2-#total_rows#");
    }
    
    // Add Freeze Pane
    SpreadsheetAddFreezePane(LOCAL.workbook, 0, 1);
    
    // Save Spreadsheet
    try {
      SpreadsheetWrite(workbook, ATTRIBUTES.file, 'yes');
      LOCAL.result = "OK";
    } catch(any e) { LOCAL.result = "FAIL"; }

    "CALLER.#ATTRIBUTES.variable#" = LOCAL.result;
  </cfscript>
</cffunction><!---end:engage()--->

<cfscript>
  /**
  * @description Populate Spreadsheet Object with data from JSON
  * @output No
  */
  public any function populateFromJSON(required workbook, required json){
    LOCAL.content = DeserializeJSON(ARGUMENTS.json);
  
    // ========== POPULATE HEADER ==============================================
    for(
      LOCAL.col = 1;
      LOCAL.col LTE ArrayLen(LOCAL.content.columns);
      LOCAL.col += 1
    ) {
      LOCAL.column = LOCAL.content.columns[LOCAL.col];
      SpreadsheetSetCellValue( ARGUMENTS.workbook, Trim(LOCAL.column), 1, LOCAL.col );
    }
    
    /* ========== POPULATE DATA ============================================= */
    // Loop through data rows
    for (
      LOCAL.row = 1;
      LOCAL.row LTE ArrayLen(LOCAL.content.data);
      LOCAL.row += 1
    ) {
      LOCAL.datum = LOCAL.content.data[LOCAL.row];
      
      // Loop through data columns
      for (
        LOCAL.col = 1;
        LOCAL.col LTE ArrayLen(LOCAL.datum);
        LOCAL.col += 1
      ) {
        SpreadsheetSetCellValue(
          ARGUMENTS.workbook,
          Trim(LOCAL.datum[LOCAL.col]),
          LOCAL.row+1,
          LOCAL.col
        );
      }
    }  
    return ARGUMENTS.workbook;
  }//end:populateFromJSON()  
</cfscript>
