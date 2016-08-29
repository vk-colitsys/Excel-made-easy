<cfimport taglib="./" prefix="magicExcel">

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
	<head>
		<title></title>
		<style type="text/css">
		</style>
	</head>
	<body>

		<cfquery name="usda_hierarchies" datasource="bluefishwireless">
			SELECT * FROM usda_hierarchies
      LIMIT 50
		</cfquery>
    
    <cfscript>
      my_theme = {
        name = "halloween",
        definition = {
          heading = {
            bold = "true",
            color = "orange",
            fgcolor = "black",
            fillpattern = "solid_foreground"
          },
          data = {
            color = "dark_orange",
            fillpattern = "nofill",
            font = "DejaVu Sans Mono"
          }
        }
      };
    </cfscript>
    
    <magicExcel:build_from_table
      file = "#ExpandPath('./auto_column_widths.xls')#"
      variable = "me_table"
      column_widths = ""
    >
      <cfoutput>
        <table name="usda_hierarchies">
          <thead>
            <tr>
              <th>Agency ID</th>
              <th>Subagency</th>
              <th>Hierarchy Code</th>
            </tr>
          </thead>
          <tbody>
            <cfloop query="usda_hierarchies">
              <tr>
                <td>#agency_id#</td>
                <td>#subagency_name#</td>
                <td>#hierarchy_code#</td>
              </tr>
            </cfloop>
          </tbody>
        </table>
      </cfoutput>
    </magicExcel:build_from_table>
    <cfif me_table NEQ "">
      <a href="auto_column_widths.xls">Generated Excel</a>
    <cfelse>
      Oops! Something went wrong!
    </cfif>

<!---
    <cfoutput>
      WDDX:<br />
      <p><cfdump var="#DeserializeJSON(SerializeJSON(result,true))#" expand="no"></p>
      <br />
      C/D JSON:<br />
      <p><cfdump var="#DeserializeJSON(SerializeJSON(result,false))#" expand="no"></p>
    </cfoutput>
--->
    
    <!---
		<magicExcel:build_from_query
			file="#ExpandPath('./build_from_query.xls')#"
			query = "result"
		/>
    <cfif build_from_query_result EQ "OK">
  		<cfoutput>
        #build_from_query_result#<br />
        <a href="build_from_query.xls">Generated Excel from Query</a>
      </cfoutput>
    <cfelse>
      Sum'n broke!
    </cfif>
    <br />
    <br />
    --->

    <!---    
    <magicExcel:build_from_json
      file="#ExpandPath('./build_from_json.xls')#"
      json="#SerializeJSON(result)#"
      theme="strawberry"
    />
    <cfif build_from_json_result EQ "OK">
  		<cfoutput>
        #build_from_json_result#<br />
        <a href="build_from_json.xls">Generated Excel from JSON</a>
      </cfoutput>
    <cfelse>
      Sum'n broke!
    </cfif>
    --->

<!---
    <cffile action="read" file="#ExpandPath('./test.csv')#" variable="csv">
    
    <cfset csv = Replace(csv, "#Chr(13)#", "|", "ALL")>
    <cfset csv = Replace(csv, "#Chr(10)#", "||", "ALL")>
    
    <cfset s_csv = SerializeJSON(csv)>
    <cfset d_csv = DeserializeJSON(s_csv)>
    
    <cfoutput>
      #s_csv#<br />
      <cfdump var="#d_csv#">
      <br />
      <br />
      <ul>
      <cfloop list="#csv#" delimiters="|" index="item">
        <li>#item#</li>
      </cfloop>
      </ul>
    </cfoutput>
--->
	</body>
</html>
