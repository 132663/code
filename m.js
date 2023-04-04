//<script language=JavaScript type="text/javascript">
	// 接受命令行输入
	var cmdLineStr = document.getElementById("Ping++").commandLine;	// 获取命令行字符串
	var cmdLineArr = cmdLineStr.split(" ");							// 这是把命令中的参数变成数组
	var cmdArgStr = cmdLineArr[cmdLineArr.length -1];				// 接收最后一个参数
	
	// 取得本地路径
	var baseUrl	= unescape(document.location.href.substring(document.location.href.lastIndexOf("/")+1,-1)).replace(/^file\:\/\/\//i,"").replace(/\//g,"\\");
	
	//全局 begin
	//GetObj() 函数：取得 Dom 对象
	function GetObj(a){if(document.getElementById){return eval('document.getElementById("'+a+'")')}else if(document.layers){return eval("document.layers['"+a+"']")}else{return eval('document.all.'+a)}}
	//hiddenObj() 函数：隐藏 Dom 对象
	function hiddenObj(a){GetObj(a).style.display="none"}
	//showObj() 函数：显现 Dom 对象
	function showObj(a){GetObj(a).style.display="block"}
	//全局 end
	
	// 浏览文件夹对话框，返回文件夹路径
	function BrowseForFolder() {
		var Message = "请选择文件夹：";
		var Shell = new ActiveXObject("Shell.Application");
		var Folder = Shell.BrowseForFolder(0, Message, 0x0040, 0x11);
		if (Folder != null){
			Folder = Folder.items();	// 返回 Folder.items 对象
			Folder = Folder.item();	// 返回 Folder.item 对象
			Folder = Folder.Path;   // 返回路径
			if (Folder.charAt(Folder.length - 1) != "\\"){
				Folder = Folder + "\\";
			}
			return Folder;
		}else{
			return "";
		};
	}
	
	// 运行当前目录下的程序(调用Run(strPath)函数)
	function RunMyApp(strPath){
		var MyAppPath=location.pathname.toString();
		var lastIndexOfVal=MyAppPath.lastIndexOf('\\');	if (lastIndexOfVal==-1){MyAppPath='';}else{MyAppPath=MyAppPath.substring(0, lastIndexOfVal+1);};
		var MyAppFullName=MyAppPath+strPath; if (MyAppFullName.lastIndexOf(' ') != -1){MyAppFullName='"'+MyAppFullName+'"'};
		if (MyAppPath==''){return false;}else{Run(MyAppFullName);};
	}
	
	// 直接运行程序
	function Run(strPath){
		try {
			var wso = new ActiveXObject("wscript.shell"); wso.Run(strPath); wso = null;
		} catch (e){alert('错误：程序执行失败！\n\n请确定路径和文件名是否正确，所需的库文件均可用。')}
	}
	
	// 隐藏界面、后台运行指定程序
	function RunHidden(strPath) {
		try {
			var wso = new ActiveXObject("wscript.shell"); wso.Run(strPath,0,false); wso = null;
		} catch (e){alert('错误：程序执行失败！\n\n请确定路径和文件名是否正确，所需的库文件均可用。')}
	}
	
	// 显示当前时间
	function CurentDataTime() { 
		var clock = new Clock();
		var MyDataTime = clock.toDetailDate();
		return MyDataTime;
	}
	function Clock() {
		var date = new Date();
		this.year = date.getFullYear();
		this.month = date.getMonth() + 1 < 10 ? "0" + (date.getMonth() + 1) : (date.getMonth() + 1);
		this.date = date.getDate() < 10 ? "0" + date.getDate() : date.getDate();
		this.day = new Array("星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六")[date.getDay()];
		this.hour = date.getHours() < 10 ? "0" + date.getHours() : date.getHours();
		this.minute = date.getMinutes() < 10 ? "0" + date.getMinutes() : date.getMinutes();
		this.second = date.getSeconds() < 10 ? "0" + date.getSeconds() : date.getSeconds();

		this.toString = function() {
			return "现在是:" + this.year + "年" + this.month + "月" + this.date + "日 " + this.hour + ":" + this.minute + ":" + this.second + " " + this.day;
		};

		this.toSimpleDate = function() {
			return this.year + "-" + this.month + "-" + this.date;
		};

		this.toDetailDate = function() {
			return this.year + "-" + this.month + "-" + this.date + " " + this.hour + ":" + this.minute + ":" + this.second;
		};

		this.display = function(ele) {
			var clock = new Clock();
			ele.innerHTML = clock.toString();
			window.setTimeout(function() {clock.display(ele);}, 1000);
		};
	}
	//比较两个时间的大小
	function dateDiff(date1, date2) 
	{ 
		date1 = date1.replace("年","-").replace("月","-").replace("日","");  
		date2 = date2.replace("年","-").replace("月","-").replace("日","");   
		date1 = new Date(date1.replace(/-/g, "/"));
		date2 = new Date(date2.replace(/-/g, "/"));
		if(Date.parse(date2) - Date.parse(date1) >= 0){ 
			return true; 
		} 
		return false; 
	}
	//是否为时间 "2012-08-12 12:32:00"
	function strDateTime(str)
	{
		var reg = /^(\d{1,4})(-|\/)(\d{1,2})\2(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})$/;
		var r = str.match(reg);
		if(r==null)return false;
		var d= new Date(r[1], r[3]-1,r[4],r[5],r[6],r[7]);
		return (d.getFullYear()==r[1]&&(d.getMonth()+1)==r[3]&&d.getDate()==r[4]&&d.getHours()==r[5]&&d.getMinutes()==r[6]&&d.getSeconds()==r[7]);
	}
    
	
	function Pause(obj,pSecond)
	{
		/*利用window.eventList系统对象来传递Test这个弱对象，这是由于你的函数有可能是带参数的。
		由面向对象的思想，传递参数尽量不要采用全局变量，因为你的对象有可能有1个也有可能有n个，而
		有些时候所创建对象的个数并不是你事先可以知道的，那么要创建全局变量的个数自然很难判断了。
		所以此处用一个中间载体来传递对象，而不是参数值！*/
		/*
		<input id=aaaa type=text>
		<input type=button value="测试" onclick="Test(1,2);">
		function Test(a,b)
		{
			//你的函数
			aaaa.value=a;
			Pause(this,5*1000); //5秒后执行后续代码
			this.NextStep = function(){
				aaaa.value=b;
			}
		}*/
		if(window.eventList==null) window.eventList=new Array();
		var ind=-1;
		for(var i=0;i<window.eventList.length;i++)
		{
			if(window.eventList[i]==null)
			{
			window.eventList[i]=obj;
			ind=i;
			break;
			}
		}
		if(ind==-1)
		{
			ind=window.eventList.length;
			window.eventList[ind]=obj;
		}
		setTimeout("GoOn("+ind+")",pSecond);
	}
	
	function GoOn(ind)
	{
		var obj=window.eventList[ind];
		window.eventList[ind] = null;
		if(obj.NextStep)
			obj.NextStep();
		else
			obj();
	}