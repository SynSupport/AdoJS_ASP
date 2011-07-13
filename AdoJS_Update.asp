<%@ LANGUAGE = "JavaScript" %> 
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
* Source:     AdoJS_Update.asp
*                                                                             
* Facility:   Example ASP using ADO objects to perform an database update 
*               using JavaScript.
*                                                                             
* Abstract:   ADO methods and properties are utilized within the BODY
*               section of the HTML document. Although this method has
*               merit, it is far less efficient compared to the
*               method used in AdoJS_Select.asp.   
*                                                                             
*             You may change the connect string and/or SQL command            
*               as needed.                                                      
*                                                                             
* $Revision:     $                                                            
*                                                                             
* $Date:         $                                                            
*                                                                             
--------------------------------------------------------------------------    
-->
<HEAD>
    <TITLE>Update Database</TITLE>
</HEAD>

<BODY BGCOLOR="White" topmargin="10" leftmargin="10">
<!-- Display Header -->
<font size="4" face="Arial, Helvetica"><b>Simple ADO Update with ASP Using JavaScript</b>

<hr size="1" color="#000000">

<%
   var oConn;      // object for ADODB.Connection obj
   var oRs;        // object for output recordset object
   var filePath;       // Directory of authors.mdb file
   var Index;

   // Create ADO Connection Component to connect with sample database
   
   oConn = Server.CreateObject("ADODB.Connection");
   oConn.Open("DSN=xfODBC;UID=DBADMIN;PWD=MANAGER;DBQ=sodbc_sa;");
   
   // Display the current values

   oRs = oConn.Execute("SELECT in_itemid, in_name, in_price FROM public.plants WHERE {fn LCASE(in_name)} LIKE 'sour gum%'");
%>

   <font size="4" face="Arial, Helvetica">
   <b>Increasing the price of Sour Gum by 10%<BR><BR>Current Price:</b>
   </font><br>

   <TABLE border = 1>
<%  
   while (!oRs.eof) 
   { %>
      <tr>
         <% for(Index=0; Index < (oRs.fields.count); Index++) { %>
            <TD VAlign=top><% = oRs(Index)%></TD>
         <% } %>
      </tr>

      <% oRs.MoveNext();
      } 
%>

   </TABLE>
   
<%
   // Insert the price by 10%
   //
   // NOTE: To add, delete and update  recordset, it is recommended to use
   // direct SQL statement instead of ADO methods.
   
   oConn.Execute ("UPDATE public.plants SET in_price = in_price + (in_price * .10) WHERE {fn LCASE(in_name)} LIKE 'sour gum%'");
   oRs = oConn.Execute("SELECT in_itemid, in_name, in_price FROM public.plants WHERE {fn LCASE(in_name)} LIKE 'sour gum%'");
%>
    
   <BR>
   <hr shade>
   <BR>
   
   <font size="4" face="Arial, Helvetica">
   <b>New Price:</b>
   </font><br>
   
   <TABLE border = 1>
<%  
      while (!oRs.eof) 
      { %>
         <tr>
            <% for(Index=0; Index < (oRs.fields.count); Index++) { %>
               <TD VAlign=top><% = oRs(Index)%></TD>
            <% } %>
         </tr>

         <% oRs.MoveNext();
      } 
%>

   </TABLE>

<%   
   // Release resources by closing the result-set and connection
   oRs.close();
   oConn.close();
%>
   
</BODY>
</HTML>