<cfscript>
  /**
  * @description Theme the worksheet
  * @output Yes
  */
  public struct function themeWorksheet(struct doc_struct, table){
    //doc = StructCopy(doc_struct);
    doc = ARGUMENTS.doc_struct;
    

    theme = getFormatForTheme(doc.config.theme);
    if (isStruct(doc.config.custom_theme)){
      theme = doc.config.custom_theme;
    }
    
    /*
      Formatting overwrites previous format in cell/row/column.
      TIP: Format from Generic to Specific
    */
    if (fn_args.theme_data_cells) {
      SpreadsheetFormatRows(doc.workbook, theme.definition.data, "1-#doc.last_data_row#");
    }

    // Explicit Column Formatting
    doc = setExplicitColumnFormatting(doc, ARGUMENTS.table);

    // Explicit Row Formatting
    doc = setExplicitRowFormatting(doc, ARGUMENTS.table);

    // FORMAT Row Headings
    if (doc.data_origin.column GT 1 AND fn_args.theme_row_headings) {
      SpreadsheetFormatColumns(doc.workbook, theme.definition.heading, "1-#doc.data_origin.column-1#");
    }

    // FORMAT Column Headings
    if (doc.data_origin.row GT 1 AND fn_args.theme_column_headings) {
      SpreadsheetFormatRows(doc.workbook, theme.definition.heading, "1-#doc.data_origin.row-1#");
    }

    // FORMAT Footer
    if (doc.footer_rows GT 0 AND fn_args.theme_footer) {
      SpreadsheetFormatRows(doc.workbook, theme.definition.heading, "#doc.footer_start_row#-#doc.footer_start_row+doc.footer_rows#");
    }

    return doc;
  }//end:themeWorksheet()


  /**
  * @description Get built in theme formatting
  * @output No
  */
  public struct function getFormatForTheme(required string name){
    switch(ARGUMENTS.name){
      /* ---------- BLUEBERRY ----------------------------------------------- */
      case "blueberry":
        format = {
          name = "blueberry",
          definition = {
            heading = { bold = "true", color = "light_cornflower_blue", fgcolor = "dark_blue", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- LEMON --------------------------------------------------- */
      case "lemon":
        format = {
          name = "lemon",
          definition = {
            heading = { bold = "true", color = "green", fgcolor = "light_yellow", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- STRAWBERRY ---------------------------------------------- */
      case "strawberry":
        format = {
          name = "strawberry",
          definition = {
            heading = { bold = "true", color = "dark_red", fgcolor = "rose", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- CHERRY -------------------------------------------------- */
      case "cherry":
        format = {
          name = "cherry",
          definition = {
            heading = { bold = "true", color = "white", fgcolor = "dark_red", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- LIME ---------------------------------------------------- */
      case "lime":
        format = {
          name = "lime",
          definition = {
            heading = { bold = "true", color = "bright_green", fgcolor = "olive_green", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- ORANGE -------------------------------------------------- */
      case "orange":
        format = {
          name = "orange",
          definition = {
            heading = { bold = "true", color = "grey_80_percent", fgcolor = "light_orange", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- LICORICE ------------------------------------------------ */
      case "licorice":
        format = {
          name = "licorice",
          definition = {
            heading = { bold = "true", color = "black", fgcolor = "grey_25_percent", fillpattern = "solid_foreground" },
            data = { color = "grey_50_percent", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- WHITE ON BLACK ------------------------------------------ */
      case "white_on_black":
        format = {
          name = "white_on_black",
          definition = {
            heading = { bold = "true", color = "white", fgcolor = "black", fillpattern = "solid_foreground" },
            data = { color = "black", fillpattern = "nofill" }
          }
        };
        break;
      /* ---------- DEFAULT (BLACK ON WHITE) -------------------------------- */
      case "black_on_white":
      default:
        format = {
          name = "black_on_white",
          definition = {
            heading = { color = "black", fillpattern = "nofill", bold = "true" },
            data =    { color = "black", fillpattern = "nofill" }
          }
        };
        break;
    }//end:switch()

    fd = format.definition;

    fd.heading.alignment = "left";

    // Default font to a sans-serif font
    fd.heading.font = "Arial";
    fd.data.font = "Arial";
    
    // Default font size to 10pt
    fd.heading.fontsize = 10;
    fd.data.fontsize = 10;

    // preserve ALL data entered into cell
    // This does not prevent missing leading zeros in numeric cells
    fd.heading.dataformat = "@";
    fd.data.dataformat = "@";

    // Align everything to top of respective cell
    fd.heading.verticalalignment = "VERTICAL_TOP";
    fd.data.verticalalignment = "VERTICAL_TOP";

    return format;
  }//end:formatForTheme()
</cfscript>
