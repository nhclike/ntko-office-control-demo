<%@ page language="java" contentType="text/html; charset=UTF-8" pageEncoding="UTF-8"%>
<!DOCTYPE html>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
	<title>office示例</title>
	<script src="http://cdn.bootcss.com/jquery/1.9.0/jquery.js"></script>


</head>
<body>



<div>
11111111111111
</div>
<div id="doc" style="width:1000px; height:800px;"></div>


</body>
<script type="text/javascript">


  fnBodyLoad();

  function fnBodyLoad () {
	console.log("test");
    var docObj = '<object id="TANGER_OCX" classid="clsid:C9BC4E1F-4248-4a3c-8A49-63A7D317F404"'
      + ' codebase="${pageContext.request.contextPath}/plugins/officecontrol/NtkoOfficeControlSetup.msi" width="100%" height="100%">'
      + '<param name="ProductCaption" value="浙江南北联合">'
      + '<param name="ProductKey" value="1061CFD1C3CADD35DA08D816866744EB45BB13DD">'
      + '<param name="IsUseUTF8URL" value="-1">'
      + '<param name="IsUseUTF8Data" value="-1">'
      + '<param name="BorderStyle" value="1">'
      + '<param name="BorderColor" value="14402205">'
      + '<param name="TitlebarColor" value="15658734">'
      + '<param name="TitlebarTextColor" value="0">'
      + '<param name="Titlebar" value="false">'
      + '<param name="MenubarColor" value="14402205">'
      + '<param name="MenuButtonColor" VALUE="16180947">'
      + '<param name="MenuBarStyle" value="3">'
      + '<param name="MenuButtonStyle" value="7">'
      + '<param name="WebUserName" value="NTKO">'
      + '<param name="Statusbar" value="false">'
      + '<param name="Caption" value="NTKO OFFICE文档控件示例演示 http://www.ntko.com">'
      + '<span style="color:red">不能装载文档控件。请在检查浏览器的选项中检查浏览器的安全设置。</span>'
      + '</object>';
    $("#doc").append(docObj);
    //获取TANGER_OCX对象
    TANGER_OCX_OBJ = document.all('TANGER_OCX');
	console.log("TANGER_OCX_OBJ:"+TANGER_OCX_OBJ);
    //如果网络错误，不弹出提示
    TANGER_OCX_OBJ.IsShowNetErrorMsg = false;
    TANGER_OCX_OBJ.FileNew = false;
    TANGER_OCX_OBJ.FileOpen = false;
    TANGER_OCX_OBJ.FileClose = false;
    TANGER_OCX_OBJ.IsShowFullScreenButton = true;

    //如果没有打开文件，则创建文件
    if (TANGER_OCX_OBJ.ActiveDocument == 'undefined' || TANGER_OCX_OBJ.ActiveDocument == null) {
      //TANGER_OCX_OBJ.CreateNew("Word.Document");
        //TANGER_OCX_OBJ.OpenFromURL(url, false, "Word.Document");
      TANGER_OCX_OBJ.OpenLocalFile('C:/Users/lij/Desktop/testfile/test.docx', false,"Word.Document");
        TANGER_OCX_OBJ.SetReadOnly=false;
      console.log(TANGER_OCX_OBJ.FileNew);
    }
  }


	

</script>
</html>