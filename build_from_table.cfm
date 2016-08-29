<cfscript>
	/*
		AUTHOR: Ryan Johnson (rhino.citguy@gmail.com)
		LICENSE: GPLv3 (http://www.gnu.org/licenses/gpl.html)
		VERSION:
		DESCRIPTION:
			Custom Tag to transform HTML Tables into a Microsoft Excel document.
		TODO:
			- Custom Theming via JSON?
			- Explicit Column Widths via column styling
	*/
	ThisTag.name = "BUILD_FROM_TABLE";

	// Tag Options Param'ed
	param ATTRIBUTES.auto_freeze_headers = "false";
	param ATTRIBUTES.file = "";
	param ATTRIBUTES.theme = "black_on_white";
  param ATTRIBUTES.theme_data_cells = "false";
	param ATTRIBUTES.theme_footer = "false";
	param ATTRIBUTES.theme_column_header = "false";
	param ATTRIBUTES.theme_row_header = "false";
	param ATTRIBUTES.freeze_at = "1,1";
	param ATTRIBUTES.custom_theme = "";
  param ATTRIBUTES.column_widths = "auto";

	// variable name to set in CALLER scope
	param ATTRIBUTES.variable = "htmltables";

	// Various Checks prior to executing the code
	if (ATTRIBUTES.file EQ "") { throw(message="FILE is required for #ThisTag.name#"); }

	/*
		This tag should only execute code upon successful completion of the
		caller end tag.
	*/
	if (ThisTag.hasEndTag){
		if (ThisTag.ExecutionMode EQ "start") {
		} else {
			include "_functions.cfm";
			engage();
		}
	} else {
		throw(message="#ThisTag.name# requires a matching end tag");
	}
</cfscript>

<!--- ========== PRIMARY TAG FUNCTION ===================================== --->
<cffunction
	name="engage"
	output="yes"
	returntype="any"
	description="I execute the logic for this tag."
	hint="Make it so, number one!"
>
	<cfscript>	
		LOCAL.content = ThisTag.GeneratedContent; // Capture Content
		ThisTag.GeneratedContent = ""; // Prevent browser display

		if (
			table2excel({
				tables = "#LOCAL.content#",
				auto_freeze_headers = "#ATTRIBUTES.auto_freeze_headers#",
				file_location = "#ATTRIBUTES.file#",
				theme = "#ATTRIBUTES.theme#",
				theme_footer = "#ATTRIBUTES.theme_footer#",
				theme_column_headings = "#ATTRIBUTES.theme_column_header#",
				theme_row_headings = "#ATTRIBUTES.theme_row_header#",
        theme_data_cells = "#ATTRIBUTES.theme_data_cells#",
				freeze_at = "#ATTRIBUTES.freeze_at#",
				custom_theme = "#ATTRIBUTES.custom_theme#",
        column_widths = "#ATTRIBUTES.column_widths#"
			})
		) {
			"CALLER.#ATTRIBUTES.variable#" = LOCAL.content;
		} else {
			"CALLER.#ATTRIBUTES.variable#" = "";
		}
	</cfscript>
</cffunction>
