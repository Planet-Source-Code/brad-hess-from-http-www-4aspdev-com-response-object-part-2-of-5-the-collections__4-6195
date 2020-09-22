<div align="center">

## Response Object \(Part 2 of 5\): The collections


</div>

### Description

Part II of our tour of the handy Response object in the ASP Object Model explores the collections available for your use.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brad Hess from http://www\.4aspdev\.com](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brad-hess-from-http-www-4aspdev-com.md)
**Level**          |Beginner
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |ASP \(Active Server Pages\), VbScript \(browser/client side\)

**Category**       |[ASP Server Object Model](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/asp-server-object-model__4-32.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brad-hess-from-http-www-4aspdev-com-response-object-part-2-of-5-the-collections__4-6195/archive/master.zip)





### Source Code

<font size="2">Part <a href="/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=6194">1</a>,<a href="/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=6195">2</a>,<a href="/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=6196">3</a>,<a href="/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=6197">4</a>,<a href="/vb/scripts/ShowCode.asp?lngWId=4&txtCodeId=6198">5</a></font></p>
<p><font face="Verdana">Alright we have covered all of the properties of the
Response Object so now lets look at the Collections.</font></p>
<p><font face="Verdana">The only collection in the response Object is the <b>Cookies</b>
Collection.  This collection allows you to use the HTTP response header to
write cookies on the client machine.  In an earlier example I used
JavaScript to write a cookie value to the clients machine.  The exact same
thing can be done using the cookies collection.  The cookies collection has
three properties: <b>Item</b>, <b>Key </b>and <b>Count</b>. To set the value of
a cookie you would do the following: <b>Response.Cookie("Color") =
"Red"</b>.  If you remember from the code example in the last
section a cookie can be what is called a dictionary cookie.  In a
dictionary cookie each element can be referenced by name.  I will go into
detail on how to set and use dictionary cookies later, but for now lets move on
to the attributes of the cookie.  Each element of a cookie has the
following attributes associated with it: <b>Domain</b>, <b>Expires</b>, <b>HasKeys</b>,
<b>Path</b>, and <b>Secure</b>.  </font></p>
<ul>
 <li><font face="Verdana">The <b>Domain</b> allows you to set the domain that
 the cookie is set to.  This is a write-only property.  Let's say
 that you wanted to set a cookie that was sent to 4aspdev.com every time
 someone requested a page in the 4aspdev.com domain.  The code would
 read:  <b>response.cookie("Color") = "4aspdev.com"</b>. 
 This cookie would get sent to 4aspdev.com every time the client requested a
 page on that domain. </font>
 <li><font face="Verdana">The <b>Expires </b>attribute sets when the cookie
 expires and is removed form the clients machine.  If this value is not
 set then the cookie is destroyed when the clients session ends. If the date
 is set to before the current date the cookie will be destroyed when the
 session ends.  To use this attribute the code would read <b>Response.Cookies("Color").Expires
 = #1/31/2000#</b>. (note the use of # around the date) </font>
 <li><font face="Verdana">The <b>HasKeys</b> attribute is used to determine if
 the cookie is a dictionary cookie and if it has subkeys.  I will go
 into this more later.  </font>
 <li><font face="Verdana">The <b>Path</b> attribute sets the path of the
 virtual directory on the server where cookies are sent. Lets say we only
 wanted cookies set to the /colors/personal virtual directory the code would
 read <b>response.cookies("Color") = "/colors/personal"</b>. 
 This sets it so that the cookie is only sent to this directory.  </font>
 <li><font face="Verdana">The<b> Secure</b> property allows you to specify
 whether the cookie is sent to the server only when the client is using
 Secure Socket Layer. This is a true False value and is default to False To
 set the value the code would read: <b>response.Cookie("color").secure
 = True</b>.</font></li>
</ul>

