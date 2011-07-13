<%@ LANGUAGE = "JavaScript1.3" %> 
<%  
    Response.Expires = -1; // force cache expiration to ensure reload 
    Response.Buffer = true; // buffer server-side task until completion (IIS5 default)
%>

<HTML>
<!--
* ----------------------------------------------------------------------      
*                                                                             
*                  Synergy - Synergy Language Version 7                       
*                                                                             
*                            Copyright (C) 2001
*     by Synergex International Corporation.  All rights reserved.            
*                                                                             
*         May not be copied or disclosed without the permission of            
*                 Synergex International Corporation                          
*                                                                             
* -----------------------------------------------------------------------     
* -----------------------------------------------------------------------     
*                                                                             
* Source:     AdoJS_Select.asp
*                                                                             
* Facility:   Example for utilizing Microsoft Universal Data Access components
*               from ASP using JavaScript.
*                                                                             
* Abstract:   Opens a connection to the PLANTS sample database. Using Advanced
*               Data Objects (ADO), issue a query returning in_itemid,
*               in_name, in_price columns from the plants database. The 
*               recordset resulting from the query is transmitted to the
*               client within an HTML TABLE tag. The HTTP header is set to
*               expire immediately (forcing a reload) and buffer all
*               server-side operations until complete (increasing through-put).
*                                                                             
*             Change the connect string (strConnect) and/or SQL 
*               command (StrQuery) as needed.
*                                                                             
* $Revision:     $                                                            
*                                                                             
* $Date:         $                                                            
*                                                                             
--------------------------------------------------------------------------    
-->

<HEAD>
<TITLE>Simple ADO Query with ASP Using JavaScript</TITLE>
</HEAD>

<BODY BGCOLOR="White" topmargin="10" leftmargin="10">

    <!-- Display Header -->
    
    <font size="4" face="Arial, Helvetica">
    <b>Simple ADO Query with ASP Using JavaScript</b>
    </font><br>
    
    <hr size="1" color="#000000">
    
    List of Available Plants:<br><br>
    
<%
    var oConn;	// connection object also containing connection.errors objects
    var oErr; // Error object collection
    var oRs; // Recordset object - 0, 1, or n recordset rows.
    var ix;
    var iy;
    var adStateClosed = 0; // From ADOVBS.INC
    var strConnect = new String("DSN=xfODBC;UID=DBADMIN;PWD=MANAGER;DBQ=sodbc_sa;");
    var strQuery = new String("SELECT in_itemid, in_name, in_price FROM plants");
    
    try 
    {
        // Instantiate the ADODB objects
        
        oConn = Server.CreateObject("ADODB.Connection");
        oRs = Server.CreateObject("ADODB.Recordset");
        oErr = Server.CreateObject("ADODB.Error");

        // Open a connection, clear any warning messages
        
        oConn.Open(strConnect);
        oConn.Errors.Clear();

        // Issue a SQL query, creating a recordset object

        oRs = oConn.Execute(strQuery);

        // Populate an HTML table with the recordset object data 
        
        Response.Write("<TABLE border = 1><br>");
  
        while (!oRs.eof) 
        { 
            Response.Write("<tr>");
      
            for(ix = 0; ix < (oRs.fields.count); ix++) 
            { 
                Response.Write("<TD VAlign=top>");
                Response.Write(oRs(ix));
                Response.Write("</TD>");
            } 

            Response.Write("</tr>");
            oRs.MoveNext();
        } 

        Response.Write("</TABLE>");
    }
    catch(e)
    {
         // Clear the response buffer creating a HTML page with only the error message
         
         Response.Clear();
         Response.Write("<html><head></head><body><h1>An application error occurred</h1><br><hr size=5>");
         
         // Dump contents of the Error object
         
         Response.Write("Error # " + e.number + " - " + e.description + "<br><hr>");
         
         if (oConn.Errors.Count > 0) 
         {
            for (var i = 1; i < oConn.Errors.Count; i++) 
            {
               oErr = oConn.Errors(i);
               strError = "Connection Error #" & oErr.Number + "<br>" +
               "   " + oErr.Description + "<br>" +
               "   (Source: " & oErr.Source & ")" + "<br>" +
               "   (SQL State: " & oErr.SQLState + ")" + "<br>" +
               "   (NativeError: " & oErr.NativeError + ")" + "<br>";
               if (oErr.HelpFile == "")
               strError = strError +
               "   No Help file available" +
               "<br><br>";
               else
               strError = strError +
               "   (HelpFile: " & oErr.HelpFile & ")" & "<br>" +
               "   (HelpContext: " & oErr.HelpContext & ")" +
               "<br><br>";
               Response.Write("<p>" & strError & "</p>");
            }
         }
         
         // Clear the error container
         oConn.Errors.Clear();
    }
    finally
    {
        // Close the recordset and/or connection. This will always be executed.
                
        if (oRs.State != adStateClosed) oRs.Close();
        if (oRs.State != adStateClosed) oConn.Close();
    }
%>

</BODY>
</HTML>
