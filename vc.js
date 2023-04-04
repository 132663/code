//<script language=JavaScript type="text/javascript">
	function GetSystemVersion() {
		var os = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem");
		for (var e = new Enumerator(os); ! e.atEnd(); e.moveNext()) {
			var v = e.item().Version;
			var ss = v.split('.');
			return ss[0] + ss[1];
		}
		return - 1;
	}
	//if (GetSystemVersion() >= 60) {
		// var cmd = location.pathname;
		// if (cmd.substring(cmd.length - 4) != ".HTA") {
			// var Shell = new ActiveXObject("Shell.Application");
			// Shell.ShellExecute("mshta.exe", "\"" + cmd.substring(0, cmd.length - 4) + ".HTA\"", "", "runas", 1);
			// window.close();
			// exit(0);
		// }
	// }
//</script>
//<script type="text/javascript">
	// 改变窗体大小
	//var windowW=400;	//窗体宽
	//var windowH=230;	//窗体高
	//var canresize = false;//是否可以改变大小
	//window.resizeTo(0,0);
	//window.moveTo((screen.width-windowW)/2,(screen.height-windowH)/2);
	//window.resizeTo(windowW,windowH);
	//window.onresize=function(){if(!canresize)window.resizeTo(windowW,windowH);}
	
	// 错误提示
	// window.onerror = reportError;
	// function reportError(msg,url,line) {
		// var str = "You have found an error as below: \n\n";
		// str += "Err: " + msg + " on line: " + line;
		// alert(str);
		// return true;
	// }
	
	// 设置变量
	
	var RunOnceView = true;	// 是否为第一次运行
	var baseUrl = unescape(document.location.href.substring(document.location.href.lastIndexOf("/")+1,-1)).replace(/^file\:\/\/\//i,"").replace(/\//g,"\\");
	var baseName = unescape(document.location.href.substring(document.location.href.lastIndexOf("/")+1,document.location.href.length));
	var MeAppPath = baseUrl + baseName;	// 自身完整路径
	var areaArr = new Array();			// 所有区域划分
	var hostArr = new Array();			// 指定区域的所有电脑
	var infoArr = new Array();			// 指定电脑的所有状态

	// 是否为查错模式
	var isDebug = true;
	
	// 提取配置文件内容
	var MeAppCfg = baseUrl + "check_IP.ini";	//"E:\\Remote Shutdown.ini";	// 配置文件路径
	//var MeAppLog = baseUrl + baseName + ".log";	//"E:\\Remote Shutdown.log";	// 日志文件路径
	
	var windowW=800;	//窗体宽
	var windowH=600;	//窗体高
	var canresize=false;//是否可以改变大小
	var windowW2=windowW;//编辑时窗体宽
	var windowH2=200;//编辑时窗体高
	window.resizeTo(0,0);
	window.moveTo((screen.width-windowW)/2,(screen.height-windowH)/2);
	window.resizeTo(windowW,windowH);
	
	// 窗口最大化与还原
	var winWidth=0;
	var winHeight=0;
	function findDimensions() //函数：获取尺寸
	{
		//获取窗口宽度
		if(window.innerWidth)
			winWidth=window.innerWidth;
		else if((document.body)&&(document.body.clientWidth))
			winWidth=document.body.clientWidth;
		//获取窗口高度
		if(window.innerHeight)
			winHeight=window.innerHeight;
		else if((document.body)&&(document.body.clientHeight))
		winHeight=document.body.clientHeight;
		/*nasty hack to deal with doctype swith in IE*/
		//通过深入Document内部对body进行检测，获取窗口大小
		if(document.documentElement && document.documentElement.clientHeight && document.documentElement.clientWidth)
		{
			winHeight=document.documentElement.clientHeight;
			winWidth=document.documentElement.clientWidth;
		}
		if (!(winWidth == windowW) || !(winHeight == windowH)){
			window.resizeTo(0,0);
			window.moveTo((screen.width-windowW)/2,(screen.height-windowH)/2);
			window.resizeTo(windowW,windowH);
		}
	}
	var winSta="normal";
	function changeWinSta(){
		if (winSta=="normal") {
			parent.moveTo(0,0);parent.resizeTo(screen.availWidth,screen.availHeight);winSta='max'
		}else{
			findDimensions();winSta="normal";
		};
	}
	
	// 移动窗口
	function move_title()
	{
		var moveing=false,move_error=false,x,y;
		titbar=document.getElementById("TITLE_BAR");
		window.onblur=function(){moveing=false;}
		titbar.onmousedown=function(){
			if(event.button==1){
				x=event.x;
				y=event.y;
				event.srcElement.setCapture();
				moveing=true;
			}
		}
		titbar.onmousemove=function(){
			if(moveing){
				try{
					window.moveBy(event.x-x,event.y-y);
				}catch(e){
					move_error=true;
				}
			}
		}
		titbar.onmouseup=function(){
			event.srcElement.releaseCapture();
			moveing=false;
			if(move_error){
				try{
					window.moveBy(event.x-x,event.y-y);
				}catch(e){}
			}
		}
	}
	
	
    // <body onkeydown="KeyDown()" oncontextmenu="event.returnValue=false" onhelp="Showhelp();return false;" scroll=no topmargin=0 leftmargin=0>
    function KeyDown(){	//屏蔽鼠标右键、Ctrl+n、shift+F10、F5刷新、退格键
        //alert("ASCII代码是："+event.keyCode);
        if ((window.event.altKey)&&
            ((window.event.keyCode==37)||   //屏蔽 Alt+ 方向键 ←
            (window.event.keyCode==39))){  //屏蔽 Alt+ 方向键 →
            alert("不准你使用ALT+方向键前进或后退网页！");
            event.returnValue=false;
        }
        if ((event.keyCode==8)  ||                 //屏蔽退格删除键
            (event.keyCode==116)||                 //屏蔽 F5 刷新键
            (event.keyCode==112)||                 //屏蔽 F1 刷新键
            (event.ctrlKey && event.keyCode==82)){ //Ctrl + R
            event.keyCode=0;
            event.returnValue=false;
        }
        if ((event.ctrlKey)&&(event.keyCode==78))   //屏蔽 Ctrl+n
            event.returnValue=false;
        if ((event.shiftKey)&&(event.keyCode==121)) //屏蔽 shift+F10
            event.returnValue=false;
        if (window.event.srcElement.tagName == "A" && window.event.shiftKey) 
            window.event.returnValue = false;  //屏蔽 shift 加鼠标左键新开一网页
        if ((window.event.altKey)&&(window.event.keyCode==115)){ //屏蔽Alt+F4
            window.showModelessDialog("about:blank","","dialogWidth:1px;dialogheight:1px");
            return false;}
    }
	
    function Showhelp(){
        alert("呵呵");
        return false;
    }
	
	
	// 程序完全载入后执行
	window.onload=function(){
	
		// 检测是否重复运行
		var MeRunCount = IsRun("mshta.exe", MeAppPath);
		if (MeRunCount >= 2){
			alert("警告：\n\n程序已运行，请不要重复运行本程序！！");
			window.close();
		};
		
		// 监控移动标题烂
		move_title()
		
	
		// 延时执行主程序，防止程序假死
		window.setTimeout(function(){main();}, 1*100);		// 延时
		
		// 处理快捷键事件
		document.onkeydown=function(){
		
			// 按键的 keyCode
			var e=event || window.event; var keyCode=e.keyCode || e.which;
			
			// 按 ESC 键 关闭窗口
			if (keyCode==27){if (confirm('确定退出？')){quit();window.close();};};
			
			// ctrl + z 锁定当前用户登陆状态
			if (e.ctrlKey&&e.keyCode==90){RunHidden("\"%windir%\\System32\\rundll32.exe\" user32.dll,LockWorkStation");};
			
			// ctrl + s 打开配置文件
			if (e.ctrlKey&&e.keyCode==83){showCfg();};

			
			// 禁止 F5 键 刷新
			//if (keyCode==116){return false;}; 
			if (keyCode==116){window.onbeforeunload();};
			
		};
	
	}
	
	// 主程序
	function main(){
    
        // 清理自身启动的脚本防止重复调用
        clear_app();
	
		// 初始化配置文件
		if (Exist(MeAppCfg) != true) {
			var str = "; 程序配置" + "\r\n" +
						"; 1. 方括号內为区域名称 " + "\r\n" +
						"; 2. 电脑列表请使用|符号分割" + "\r\n" + "\r\n" +
						"; 示例：" + "\r\n" + "\r\n" + "\r\n" +
						"[XX公司 - A部門]" + "\r\n" +
						"电脑列表 = kd-010 | kd-012 | kd-032 | kd-065 | kd-036 | kd-136 | kd-214 | kd-235 | kd-089 | kd-097 | kd-054 | kd-086 | kd-067" + "\r\n" +
						"[XX公司 - B部門]" + "\r\n" +
						"电脑列表 = CD-010 | cd-012 | cd-032 | " + "\r\n"
			WriteTextUnicode(str, MeAppCfg);
			alert(str);
		}
        
        // 清理注册表
        var strKey = "HKCU\\Software\\Ping++\\";
        var Shell = new ActiveXObject("WScript.Shell");
        try 
        {
            //Shell.RegDelete(strKey);
            Shell.run('cmd /c reg delete ' + strKey + ' /f', 0, true);
        } catch (e){};
        Shell = null;
		
		// 初始化配置文件
		init_conf();
	
		// 界面部分
		if (cmdArgStr == ""){
		
			// 初始化界面
			init_list();
			
			// 启动检测程序
			start_check_state();
			
			// 获取状态
			get_host_state();
			
			// 显示状态
			show_host_state();
		}
	}

	function init_conf(){
		// 所有的区域(数组)
		var src = read_unicode_file(MeAppCfg);		// 读取配置文件
		var re = /\[.+\]/g;							// 创建正则表达式模式。
		var arr;
		while ((arr = re.exec(src)) != null){
			areaArr[areaArr.length] = arr.toString().replace(/[\[\]]/g, "");			// 区域数组
		}
		
		// 区域的所有电脑
		for (var i = 0; i < areaArr.length; i++)
		{
			var mySection = areaArr[i];
			hostArr[mySection] = ReadIniUnicode( MeAppCfg, mySection, "电脑列表" ).toString().replace(/\s*/g, "").split("|");
		}
		
		// 电脑的所有状态
		for (var i=0; i<areaArr.length; i++)
		{
			for (var j = 0; j < hostArr[areaArr[i]].length; j++)
			{
				var host = hostArr[areaArr[i]][j].toString().replace(/\(.*\)/g, "");
				var host_name = "", arr = "";
				var src = hostArr[areaArr[i]][j];
				var re = /\(.*?\)/g;							// 创建正则表达式模式。
				while ((arr = re.exec(src)) != null){
					host_name = arr.toString().replace(/[\(\)]/g, "");			// 区域数组
				}
				if (host_name == ""){host_name = host;}else{host_name += " " + host;};
				
				// 初始化电脑所有状态信息
				infoArr[host] = new Array();
				infoArr[host]["host_name"] = host_name;
				infoArr[host]["Package_Duration"] = "unkown";
				infoArr[host]["Package_DurationStr"] = "unkown";
				infoArr[host]["Package_LastCheckTime"] = "unkown";
				infoArr[host]["Package_Lost"] = "unkown";
				infoArr[host]["Package_Sent"] = "unkown";
				infoArr[host]["Package_StatusCode"] = "unkown";
				infoArr[host]["Reply_bytes"] = "unkown";
				infoArr[host]["Reply_time"] = "unkown";
				infoArr[host]["Reply_time"] = "unkown";
				infoArr[host]["Reply_time_Average"] = "unkown";
				infoArr[host]["Reply_time_Count"] = "unkown";
				infoArr[host]["Reply_time_Max"] = "unkown";
				infoArr[host]["Reply_time_Min"] = "unkown";
				infoArr[host]["Reply_TTL"] = "unkown";
			}
		}
	}
	
	// 初始化界面
	function init_list(){
		var doc_title;
		for (var i=0; i<areaArr.length; i++)
		{
			doc_title = areaArr[0];
			var info_html = "", host_html = "";
			host_html += '<br><br>\r\n<span class="area_name" contenteditable>' + areaArr[i] + '</span><br>\r\n';
			host_html += '<table width="100%" border="1" borderColor=#EEEEEE  cellspacing="1" cellPadding="1" style="BORDER-COLLAPSE: collapse">\r\n<tbody>\r\n<tr>\r\n';
			for (var j = 0; j < hostArr[areaArr[i]].length; j++) {
				// 网络地址
				var host = hostArr[areaArr[i]][j].toString().replace(/(\(.*\))?/g, "");
				// 获取备注，如 “baidu.com (百度)” 中的“百度”
				var host_name = "", arr = "";
				var src = hostArr[areaArr[i]][j];
				var re = /\(.*?\)/g;							// 创建正则表达式模式。
				while ((arr = re.exec(src)) != null){
					host_name = arr.toString().replace(/[\(\)]/g, "");			// 区域数组
				}
                if (host_name == ""){host_name = host;};
				
				// 一行 10 个单元格，满了则切换到下一行
				if ((j != 0) && (j%10 == 0)) {host_html += '</tr>\r\n<tr>\r\n';};
				
				// 配置文件中是否存在重复的网络地址
				if (GetObj(host) == null){
					host_html += '<td width="10%" class="tdc2">\r\n'+
								'<div class=HostName id=' + host + ' name=' + host + ' contenteditable onclick=\'showObj("info_' + host + '");GetObj("info_' + host + '").focus();this.style.display="none";\'>' + host_name + '</div>\r\n'+
								'<div class=HostData id="info_' + host + '" name="info_' + host + '" contenteditable style="display=none;" oncontextmenu=\'this.style.display="none";showObj("' + host + '");GetObj("' + host + '").focus()\' title="按鼠标右键隐藏"></div>\r\n' +
								'</td>\r\n';
				} else {
					host_html += '<td width="10%" class="tdc2">\r\n'+
								'<div class=RepeatHostName onclick=\'GetObj("' + host + '").click();\' title="该网络地址与上文重复">' + host_name + '</div>\r\n'+
								'</td>\r\n';
				}
			}
			host_html += '</tr>\r\n</tbody>\r\n</table>\r\n';
			
			// 将数据输出到指定标签
			GetObj("check_list").innerHTML = GetObj("check_list").innerHTML + host_html;
		}
		
		// 将配置文件中第一个节点名加入程序标题
		document.title += " (" + doc_title + ")";
	}
	
	
	// 启动检查程序
	function start_check_state()
	{
		for (var i=0; i<areaArr.length; i++)
		{
			var mySectionNum = i + 1;
			if (isDebug == true)
			{
				var script_app_path = baseUrl + "check_state.vbe";
				RunNotWait( 'cmd.exe /c cscript.exe //nologo "' + script_app_path + '" ' + mySectionNum + " & echo ^-^> Quit & ping -n 10 127.1>nul");
			} else {
				var script_app_path = baseUrl + "check_state.vbe";
				RunHideNotWait( 'wscript.exe "' + script_app_path + '" ' + mySectionNum);
			}
		}
	}
	
	// 获取状态信息
	function get_host_state()
	{
		for (var i=0; i<areaArr.length; i++)
		{
			for (var j = 0; j < hostArr[areaArr[i]].length; j++) {
				var shell = new ActiveXObject('WScript.Shell');
				var host = hostArr[areaArr[i]][j].toString().replace(/\(.*\)/g, "");
				var host_name = "", arr = "";
				var src = hostArr[areaArr[i]][j];
				var re = /\(.*?\)/g;							// 创建正则表达式模式。
				while ((arr = re.exec(src)) != null){
					host_name = arr.toString().replace(/[\(\)]/g, "");			// 区域数组
				}
                if (host_name == ""){host_name = host;}else{host_name += " " + host;};
				if ((host != null) && (host != "") && (GetObj( host ) != null ))
				{
					var Package_Sent = 0;
					var Package_Received = 0;
					var Package_Lost = 0;
					var Package_StatusCode = 1;
					var Package_Duration = "";
					var Package_DurationStr = "";
					var Package_LastCheckTime = "";
					var title, info, color, border;
					var strKey = "HKCU\\Software\\Ping++\\" + host + "\\";
					
					try {Package_StatusCode = shell.RegRead( strKey + "Package_StatusCode");}catch(e){};
					try {Package_Sent = shell.RegRead( strKey + "Package_Sent");}catch(e){};
					try {Package_Duration = shell.RegRead( strKey + "Package_Duration");}catch(e){};
					try {Package_LastCheckTime = shell.RegRead( strKey + "Package_LastCheckTime");}catch(e){};
					try {Package_Lost = shell.RegRead( strKey + "Package_Lost");}catch(e){};
					
					Package_DurationStr = getDateDiff(Package_Duration);
                    var Package_Lost_val;
					Package_Lost_val = Math.round(Package_Lost * 100 / Package_Sent);
					
					// 如果检测更新时间超时，说明没有启动检测功能
					if (CheckTimeOverstep(Package_LastCheckTime) == true){Package_StatusCode = "";};

					var Reply_bytes = 0;
					var Reply_time = 0;
					var Reply_TTL = 0;
					var Reply_time_Min = 0;
					var Reply_time_Max = 0;
					var Reply_time_Count = 0;
					var Reply_time_Average = 0;
					
					try {Reply_bytes = shell.RegRead( strKey + "Reply_bytes");}catch(e){};
					try {Reply_time = shell.RegRead( strKey + "Reply_time");}catch(e){};
					try {Reply_TTL = shell.RegRead( strKey + "Reply_TTL");}catch(e){};
					
					try {Reply_time_Min = shell.RegRead( strKey + "Reply_time_Min");}catch(e){};
					try {Reply_time_Max = shell.RegRead( strKey + "Reply_time_Max");}catch(e){};
					try {Reply_time_Count = shell.RegRead( strKey + "Reply_time_Count");}catch(e){};
					try {Reply_time_Average = shell.RegRead( strKey + "Reply_time_Average");}catch(e){};
							
					infoArr[host]["Package_Duration"] = Package_Duration;
					infoArr[host]["Package_DurationStr"] = Package_DurationStr;
					infoArr[host]["Package_LastCheckTime"] = Package_LastCheckTime;
					infoArr[host]["Package_Lost"] = Package_Lost;
					infoArr[host]["Package_Sent"] = Package_Sent;
					infoArr[host]["Package_StatusCode"] = Package_StatusCode;
					infoArr[host]["Reply_bytes"] = Reply_bytes;
					infoArr[host]["Reply_time"] = Reply_time;
					infoArr[host]["Reply_time_Average"] = Reply_time_Average;
					infoArr[host]["Reply_time_Count"] = Reply_time_Count;
					infoArr[host]["Reply_time_Max"] = Reply_time_Max;
					infoArr[host]["Reply_time_Min"] = Reply_time_Min;
					infoArr[host]["Reply_TTL"] = Reply_TTL;
				}
			}
		}
		
		// 循环执行程序，心跳时间为3秒
		window.setTimeout(function(){get_host_state();}, 3*1000);		// 延时
	}
	
	
	// 显示状态
	function show_host_state()
	{
		for (var i=0; i<areaArr.length; i++)
		{
			for (var j = 0; j < hostArr[areaArr[i]].length; j++) {
				var host = hostArr[areaArr[i]][j].toString().replace(/\(.*\)/g, "");
				var host_name = "", arr = "";
				var src = hostArr[areaArr[i]][j];
				var re = /\(.*?\)/g;							// 创建正则表达式模式。
				while ((arr = re.exec(src)) != null){
					host_name = arr.toString().replace(/[\(\)]/g, "");			// 区域数组
				}
                if (host_name == ""){host_name = host;}else{host_name += " " + host;};
				if ((host != null) && (host != "") && (GetObj( host ) != null ))
				{
					var Package_Sent = 0;
					var Package_Received = 0;
					var Package_Lost = 0;
					var Package_StatusCode = 1;
					var Package_Duration = "";
					var Package_DurationStr = "";
					var Package_LastCheckTime = "";
					var Reply_bytes = 0;
					var Reply_time = 0;
					var Reply_TTL = 0;
					var Reply_time_Min = 0;
					var Reply_time_Max = 0;
					var Reply_time_Count = 0;
					var Reply_time_Average = 0;
					var title, info, color, border;
					
					Package_Duration	=	infoArr[host]["Package_Duration"];
					Package_DurationStr	=	infoArr[host]["Package_DurationStr"];
					Package_LastCheckTime	=	infoArr[host]["Package_LastCheckTime"];
					Package_Lost	=	infoArr[host]["Package_Lost"];
					Package_Sent	=	infoArr[host]["Package_Sent"];
					Package_StatusCode	=	infoArr[host]["Package_StatusCode"];
					Reply_bytes	=	infoArr[host]["Reply_bytes"];
					Reply_time	=	infoArr[host]["Reply_time"];
					Reply_time_Average	=	infoArr[host]["Reply_time_Average"];
					Reply_time_Count	=	infoArr[host]["Reply_time_Count"];
					Reply_time_Max	=	infoArr[host]["Reply_time_Max"];
					Reply_time_Min	=	infoArr[host]["Reply_time_Min"];
					Reply_TTL	=	infoArr[host]["Reply_TTL"];

					
					Package_DurationStr = getDateDiff(Package_Duration);
                    var Package_Lost_val;
					Package_Lost_val = Math.round(Package_Lost * 100 / Package_Sent);
                    
					title =	"发送数据包: " + Package_Sent + " 个\r\n" +
							"丢失数据包: " + Package_Lost +" 个 ( " + Package_Lost_val + "% 丢失 )\r\n" +
							"更新时间: " + Package_LastCheckTime + "\r\n" + "\r\n";
					
					// 如果检测更新时间超时，说明没有启动检测功能
					if (CheckTimeOverstep(Package_LastCheckTime) == true){Package_StatusCode = "";};

					switch (Package_StatusCode) {
						case 0 :
							// 网络接通时
							title = host_name + " 在线" + "\r\n" +
									"平均延时：" + Reply_time_Average + " 毫秒" + "\r\n" +
									"持续时间：" + Package_DurationStr + "\r\n" + "\r\n" +
									title +
									"响应数据包大小: " + Reply_bytes + " 比特\r\n" +
									"响应时间: " + Reply_time + " 毫秒" + "\r\n" +
									"生存时间: " + Reply_TTL + "\r\n" +
									"最小延时: " + Reply_time_Min + " 毫秒" + "\r\n" +
									"最大延时: " + Reply_time_Max + " 毫秒" + "\r\n" +
									"总计延时: " + Reply_time_Count + " 毫秒"
									
							color = "66AA33";
							border = "1px solid #66AA33";
							break;
						case 1 :
						case 11010 :
							// 网络不通时
							title = host_name + " 离线" + "\r\n" +
                                    "持续时间：" + Package_DurationStr + "\r\n" + "\r\n" +
									title ;
							color = "CC3300";
							border = "1px solid #CC3300";
							break;
						default :
							// 无法检测时
							title = host_name + " 检测失败。\r\n若长时间检测失败，请重启程序。" + "\r\n\r\n" +
                                    "持续时间：" + Package_DurationStr + "\r\n" + "\r\n" +
									title ;
							color = "999999";
							border = "1px solid #CCCCCC";
							break;
					}
					
					// 更新信息
					GetObj( host ).title = title;
					GetObj( host ).style.color = color;
					GetObj( host ).style.border	= border;
					if (GetObj( "info_" + host ) != null ){
						GetObj( "info_" + host ).innerText = title;
                        GetObj( "info_" + host ).style.color = "666666";
                        GetObj( "info_" + host ).style.border	= border;
					};
				}
			}
		}
		
		// 循环执行程序，心跳时间为3秒
		window.setTimeout(function(){show_host_state();}, 3*1000);		// 延时
	}
	
	
	// 隐藏所有主机信息
	function hidden_all_host_info()
	{
		for (var i=0; i<areaArr.length; i++)
		{
			for (var j = 0; j < hostArr[areaArr[i]].length; j++) {
				var host = hostArr[areaArr[i]][j];
				if ((host != null) && (host != "") && (GetObj( host ) != null ))
				{
					if (GetObj( "info_" + host ) != null ){
						hiddenObj( "info_" + host );
					};
					if (GetObj( host ) != null ){
						showObj( host );
					};
				}
			}
		}
	}
	

	// 读取配置文件
	function read_unicode_file(fileStr)
    {
		var fso, wtxt;
		var ForWriting = 1;			// 'ForReading = 1 (只读不写), ForWriting = 2 (只写不读), ForAppending = 8 (在文件末尾写)
		var Create = false;			// 'Boolean 值，filename 不存在时是否创建新文件。允许创建为 True，否则为 False。默认值为 False。
		var TristateTrue = -1;		// 'TristateUseDefault = -2 (SystemDefault), TristateTrue = -1 (Unicode), TristateFalse = 0 (ASCII)
		var fso = new ActiveXObject("Scripting.filesystemobject");
		var wtxt = fso.OpenTextFile(fso.GetFile(fileStr), ForWriting, Create, TristateTrue);
		//wtxt.Write str
		var str = wtxt.ReadAll();
		wtxt.Close();
		fso = null; wtxt = null;
		return str;
	}
	
	
	// 退出程序
	function quit()
	{
        // 清理自身启动的脚本
		clear_app();
	}
    
	
    // 清理自身启动的脚本防止重复调用
	function clear_app()
	{
		var script_app_path;
		if (isDebug == true)
		{
			script_app_path = baseUrl + "bin\\check_state.vbe"
		} else {
			script_app_path = baseUrl + "bin\\check_state.vbe"
		};
		KillApp("WScript.exe", script_app_path);
		KillApp("CScript.exe", script_app_path);
	}
	
	
    // 人性化的提示持续时间
    function getDateDiff(dateTimeStamp){
        var second = 1 * 1000;
        var minute = second * 60;
        var hour = minute * 60;
        var day = hour * 24;
        var halfamonth = day * 15;
        var month = day * 30;
        var year = month * 12;
        
        dateTimeStamp = Date.parse(dateTimeStamp.replace(/-/gi,"/"));
        
        var now = new Date().getTime();
        var diffValue = now - dateTimeStamp;

        if(diffValue < 0){
        //非法操作
        //alert("结束日期不能小于开始日期！");
        }
        
        var yearC = diffValue / year;
        var monthC = diffValue / month;
        var weekC = diffValue / (7*day);
        var dayC = diffValue / day;
        var hourC = diffValue / hour;
        var minuteC = diffValue / minute;
        var secondC = diffValue / second;
        
        if(yearC>=1){
            result = parseInt(yearC) + "年";
        }
        else if(monthC>=1){
            result = parseInt(monthC) + "月";
        }
        else if(weekC>=1){
            result = parseInt(weekC) + "星期";
        }
        else if(dayC>=1){
            result = parseInt(dayC) +"天";
        }
        else if(hourC>=1){
            result = parseInt(hourC) +"小时";
        }
        else if(minuteC>=1){
            result = parseInt(minuteC) +"分钟";
        }
        else if(secondC>=1){
            result = parseInt(secondC) +"秒";
        }else{
            result = "刚刚";
        };
        return result;
    }

    // 人性化的提示持续时间
    function CheckTimeOverstep(dateTimeStamp) {
        var second = 1 * 1000;
        var minute = second * 60;
        
        dateTimeStamp = Date.parse(dateTimeStamp.replace(/-/gi,"/"));
        
        var now = new Date().getTime();
        var diffValue = now - dateTimeStamp;

        if(diffValue < 0){
			//非法操作
			//alert("结束日期不能小于开始日期！");
        }
        
        var minuteC = diffValue / minute;
        var secondC = diffValue / second;
		
		// 时间差大于5分钟
		if(minuteC>=5){
		//if(secondC>=5){
            return true;
        }else{
            return false;
        };
        
    }

//</script>