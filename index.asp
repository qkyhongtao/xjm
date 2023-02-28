<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/user_xq.asp"-->
<%
Dim user_xq
Dim user_xq_cmd
Dim user_xq_numRows

Set user_xq_cmd = Server.CreateObject ("ADODB.Command")
user_xq_cmd.ActiveConnection = MM_user_xq_STRING
user_xq_cmd.CommandText = "SELECT * FROM user ORDER BY jzsj DESC" 
user_xq_cmd.Prepared = true

Set user_xq = user_xq_cmd.Execute
user_xq_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 6
Repeat1__index = 0
user_xq_numRows = user_xq_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

Dim user_xq_total
Dim user_xq_first
Dim user_xq_last

' set the record count
user_xq_total = user_xq.RecordCount

' set the number of rows displayed on this page
If (user_xq_numRows < 0) Then
  user_xq_numRows = user_xq_total
Elseif (user_xq_numRows = 0) Then
  user_xq_numRows = 1
End If

' set the first and last displayed record
user_xq_first = 1
user_xq_last  = user_xq_first + user_xq_numRows - 1

' if we have the correct record count, check the other stats
If (user_xq_total <> -1) Then
  If (user_xq_first > user_xq_total) Then
    user_xq_first = user_xq_total
  End If
  If (user_xq_last > user_xq_total) Then
    user_xq_last = user_xq_total
  End If
  If (user_xq_numRows > user_xq_total) Then
    user_xq_numRows = user_xq_total
  End If
End If
%>
<%
Dim MM_paramName 
%>
<%
' *** Move To Record and Go To Record: declare variables

Dim MM_rs
Dim MM_rsCount
Dim MM_size
Dim MM_uniqueCol
Dim MM_offset
Dim MM_atTotal
Dim MM_paramIsDefined

Dim MM_param
Dim MM_index

Set MM_rs    = user_xq
MM_rsCount   = user_xq_total
MM_size      = user_xq_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  MM_param = Request.QueryString("index")
  If (MM_param = "") Then
    MM_param = Request.QueryString("offset")
  End If
  If (MM_param <> "") Then
    MM_offset = Int(MM_param)
  End If

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While ((Not MM_rs.EOF) And (MM_index < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
  If (MM_rs.EOF) Then 
    MM_offset = MM_index  ' set MM_offset to the last possible record
  End If

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  MM_index = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or MM_index < MM_offset + MM_size))
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = MM_index
    If (MM_size < 0 Or MM_size > MM_rsCount) Then
      MM_size = MM_rsCount
    End If
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  MM_index = 0
  While (Not MM_rs.EOF And MM_index < MM_offset)
    MM_rs.MoveNext
    MM_index = MM_index + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
user_xq_first = MM_offset + 1
user_xq_last  = MM_offset + MM_size

If (MM_rsCount <> -1) Then
  If (user_xq_first > MM_rsCount) Then
    user_xq_first = MM_rsCount
  End If
  If (user_xq_last > MM_rsCount) Then
    user_xq_last = MM_rsCount
  End If
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

Dim MM_keepMove
Dim MM_moveParam
Dim MM_moveFirst
Dim MM_moveLast
Dim MM_moveNext
Dim MM_movePrev

Dim MM_urlStr
Dim MM_paramList
Dim MM_paramIndex
Dim MM_nextParam

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 1) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    MM_paramList = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For MM_paramIndex = 0 To UBound(MM_paramList)
      MM_nextParam = Left(MM_paramList(MM_paramIndex), InStr(MM_paramList(MM_paramIndex),"=") - 1)
      If (StrComp(MM_nextParam,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & MM_paramList(MM_paramIndex)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then 
  MM_keepMove = Server.HTMLEncode(MM_keepMove) & "&"
End If

MM_urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="

MM_moveFirst = MM_urlStr & "0"
MM_moveLast  = MM_urlStr & "-1"
MM_moveNext  = MM_urlStr & CStr(MM_offset + MM_size)
If (MM_offset - MM_size < 0) Then
  MM_movePrev = MM_urlStr & "0"
Else
  MM_movePrev = MM_urlStr & CStr(MM_offset - MM_size)
End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>医疗机构诊疗登记管理</title>
<style type="text/css">
body html {
	padding:0%;
	margin-top: 0;
	margin-right: auto;
	margin-bottom: 0;
	margin-left: auto;
	}

.waibao {
	width: 100%;
	margin-top: 0px;
	margin-right: auto;
	margin-bottom: 0px;
	margin-left: auto;
	clear: both;
	background-color: #F4FFF5;
}
.bti {
	background-repeat: repeat-x;
	width: 85%;
	margin-top: 0%;
	margin-right: auto;
	margin-bottom: 0%;
	margin-left: auto;
	clear: both;
	padding: 0px;
}
.top_zt1 {
	font-size: 2.1em;
	width: 80%;
	margin-top: 0%;
	margin-right: auto;
	margin-bottom: 0%;
	margin-left: auto;
	text-align: center;
	font-weight: bolder;
	color: #F03;
	padding: 1%;
}
#form1 {
	float: right;
	width: 25%;
	clear: right;
}
.dwmc {
	width: 35%;
	color: #09F;
	font-weight: bolder;
	font-size: 1.3em;
	float: left;
	clear: right;
	letter-spacing: 0.2em;
}
.dw {
	color: #09F;
}
.jrsj {
	float: right;
	clear: right;
	font-weight: bold;
	margin-top: 0.3%;
	margin-right: 1%;
	margin-bottom: 0%;
	margin-left: 0%;
}
.cfsy {
	width: 43%;
	float: left;
	clear: right;
	font-weight: bold;
	margin: 0%;
}
.right {
	width: 85%;
	margin-top: 0.5%;
	margin-right: auto;
	margin-bottom: 0%;
	margin-left: auto;
	background-image: url(img/bj2.jpg);
	background-repeat: repeat;
	text-decoration: none;
	padding-top: 0%;
	padding-right: 自动;
	padding-bottom: 0%;
	padding-left: 自动;
}

.bg {
	width: 100%;
	margin: 0%;
	opacity:0.7;
	clear: left;
	background-repeat: repeat-x;
	padding: 0%;
}
.xgai {
	color: #FFF;
	position: relative;
	right: -90px;
}
.xgg {
	width: 10%;
	position: relative;
	top: -10%;
	right: -110%;
	float: left;
}
a{
 color:#000;
 text-decoration:none;}
</style>
</head>

<body>
<div class="waibao">
<div class="bti">
<div class="top_zt1"><a style="color:#F03"; text-decoration:none;" href="index_hanma_Chufang.asp">安溪县卫生所、个体诊所医疗机构诊疗登记</a> </div>
<div class="top">
  <form id="form1" name="form1" method="post"  action="index_soushuo.asp" >
    <label for="suinp"></label>
    <input type="text" name="suinp"  />
    <input type="submit" name="button" value="搜索" />
  </form>
 </div>
<div class="dwmc"><a  class="dw" href="index.asp">单位：谢江木西医内科诊所</a></div>
<div class="cfsy" ><a style="text-decoration:none;" class="cflj" href="index_lingszongshu.asp">零售首页</a>&nbsp;&nbsp;<a  style="text-decoration:none;" class="cflj" href="index_lings_shouye.asp">零售处方</a>&nbsp;&nbsp;&nbsp;<a  style="text-decoration:none;" class="cflj" href="index_SuoyouChufang.asp">处方首页</a>&nbsp;&nbsp;&nbsp;<a  style="text-decoration:none;" class="cflj" href="index_bingliye.asp">病历首页</a>&nbsp;&nbsp;&nbsp;<a  style="text-decoration:none;" class="cflj" href="index_zhongyao_SuoyouChufang.asp">中药处方</a></div>
<div class="jrsj">今日时间：<%= now%></div>
<div>&nbsp;</div>
<div>&nbsp;</div>
</div>
<div class="right">
  <table width="85%" border="1" cellpadding="0" cellspacing="3" id="xinxi" class="bg">
    <tr>
      <td width="15%" height="35" align="center" bgcolor="#FFFFFF">病历号</td>
      <td width="10%" height="35" align="center" bgcolor="#FFFFFF">姓名</td>
      <td width="9%" align="center" bgcolor="#FFFFFF">性别</td>
      <td width="16%" align="center" bgcolor="#FFFFFF">诊断</td>
      <td width="18%" align="center" bgcolor="#FFFFFF">就诊日期</td>
      <td colspan="3" align="center" bgcolor="#FFFFFF">操作</td>
      </tr>
    <% 
While ((Repeat1__numRows <> 0) AND (NOT user_xq.EOF)) 
%>
  <tr>
    <td height="45" align="center" bgcolor="#FFFFFF"><%= int((10000-1+1)*rnd+50)%><%=(user_xq.Fields.Item("id").Value)%><%= int((100-10+1)*rnd+50)%></td>
    <td height="45" align="center" bgcolor="#FFFFFF"><%=(user_xq.Fields.Item("xm").Value)%></td>
    <td align="center" bgcolor="#FFFFFF"><%=(user_xq.Fields.Item("xb").Value)%></td>
    <td align="center" bgcolor="#FFFFFF"><%=(user_xq.Fields.Item("zd").Value)%><%=(user_xq.Fields.Item("zd1").Value)%></td>
    <td align="center" bgcolor="#FFFFFF"><%=(user_xq.Fields.Item("jzsj").Value)%></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="index_xiangqing.asp?nid=<%=(user_xq.Fields.Item("id").Value)%>" target="_new">&nbsp;详情</a></td>
    <td width="10%" align="center" bgcolor="#FFFFFF"><a href="index_xinzen.asp">新增</a></td>
    <td width="12%" align="center" bgcolor="#FFFFFF"><p><a href="index_lingshouxinzeng.asp">零售</a></p></td>
     </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  user_xq.MoveNext()
Wend
%>
<tr class="yema_zt">
      <td height="35" colspan="8" align="center" bgcolor="#FFFFFF"><table width="443" border="0">
        <tr>
          <td width="60"><a href="<%=MM_moveFirst%>">|<<</a></td>
          <td width="90"><a href="<%=MM_movePrev%>">上一页</a></td>
          <td width="90"><a href="<%=MM_moveNext%>">下一页</a></td>
          <td width="60"><a href="<%=MM_moveLast%>">>>|</a></td>
        </tr>
    </table></td>
    </tr>
  </table>
</div>
</div>
</body>
</html>
<%
user_xq.Close()
Set user_xq = Nothing
%>
