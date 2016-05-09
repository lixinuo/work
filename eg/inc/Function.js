function $2(ObjID){
	return document.getElementById(ObjID);
}

function Ajax(TagID,UrlStr){ 
	var xmlHttp=GetXmlHttpObject();
	if (xmlHttp==null){
  		alert ("您的浏览器不支持AJAX!");
		return false;
	}
	var Url="Ajax.asp"+UrlStr+"&sid="+Math.random();
	xmlHttp.onreadystatechange=function(){
		if (xmlHttp.readyState==4){ 
			if (xmlHttp.status==200){
				$2(TagID).innerHTML=xmlHttp.responseText;
			}
		}
	}
	xmlHttp.open("GET",Url,true);
	xmlHttp.send(null);
}

function GetXmlHttpObject(){
	var XmlHttp=null;
	try{
		XmlHttp=new XMLHttpRequest(); // Firefox, Opera 8.0+, Safari
	}
	catch (e){
 	 	// Internet Explorer
  		try{
			XmlHttp=new ActiveXObject("Msxml2.XMLHTTP");
		}
  		catch (e){
			XmlHttp=new ActiveXObject("Microsoft.XMLHTTP");
		}
 	}
	return XmlHttp;
}

function DisDiv(ObjID,Val){
	if (Val==0){
		$2(ObjID).style.display="none";
	}
	else{
		$2(ObjID).style.display="block";
	}
}

function DisDivs(ObjID){
	if ($2(ObjID).style.display=="none"){
		$2(ObjID).style.display="block";
	}
	else{
		$2(ObjID).style.display="none";
	}
}

function MarqueeImage(ObjDiv,Obj1,Obj2,Direction,Speed){
    var demo1 = $2(Obj1);
    var demo2 = $2(Obj2);
    var mydiv = $2(ObjDiv);
	var Tid;
	switch(Direction){
	case "left":
		if (demo1.offsetWidth<=mydiv.offsetWidth) return;
		break;
	case "right":
		if (demo1.offsetWidth<=mydiv.offsetWidth) return;
		break;
	case "top":
		if (demo1.offsetHeight<=mydiv.offsetHeight) return;
		break;
	case "bottom":
		if (demo1.offsetHeight<=mydiv.offsetHeight) return;
		break;
	}
    demo2.innerHTML=demo1.innerHTML;
    Tid=setInterval(Marquee,Speed)
    mydiv.onmouseover=function(){clearInterval(Tid)}
    mydiv.onmouseout=function(){Tid=setInterval(Marquee,Speed)}
    
    function Marquee(){
		switch(Direction){
		case "left":
            if(mydiv.scrollLeft>=demo1.offsetWidth)
                mydiv.scrollLeft=0;
            else
                mydiv.scrollLeft++;
			break;
		case "right":
            if(mydiv.scrollLeft==0)
                mydiv.scrollLeft=demo1.offsetWidth;
            else
                mydiv.scrollLeft--;
			break;
		case "top":
			if(mydiv.scrollTop>=demo1.offsetHeight)
                mydiv.scrollTop=0;
            else
                mydiv.scrollTop++;
			break;
		case "bottom":
			if(mydiv.scrollTop==0)
                mydiv.scrollTop=demo1.offsetHeight;
            else
                mydiv.scrollTop--;
			break;
		}
    }
}

function CheckSearch(Language){
	var SearchKey=$2("search_key");
	if (SearchKey.value==""){
		if (Language=="cn"){
			alert("请输入关键词！");
		}
		else{
			alert("Please enter keywords!");
		}
		SearchKey.focus();
		return;
	}
	location.href="Products.asp?SearchKey="+escape(SearchKey.value);
}

function QQ(QQMeter,MSNMeter,SkypeMeter,WangMeter){
	var ScreenWidth=screen.width;
	var Div=document.createElement("div");
    Div.id="qq_online";
    Div.style.position="absolute";
	if (ScreenWidth>=1024+130){
		Div.style.left=((ScreenWidth-1000)/2+1000+1)+"px";
	}
	else{
    	Div.style.right="1px";
	}
    Div.style.top="160px";
    
    var Html="";
	Html+="<div id=\"qq_online_top\"><p><a href=\"javascript:void(0);\" onclick=\"$2('qq_online').style.display='none';\" title=\"关闭\"><img src=\"Images/QQ_CloseAll.gif\" /></a>Online Service</p></div>";
	Html+="<div id=\"qq_online_list\">";
	var QQName,QQNumber;
	if (QQMeter!=""){
		QQMeter=QQMeter.split("|");
		for (var i=0;i<QQMeter.length;i++){
			var QQNumber=QQMeter[i].split(",")[0];
			var QQName=QQMeter[i].split(",")[1];
			Html+="<p class=\"qq_out\" onmouseover=\"this.className='qq_over';\" onmouseout=\"this.className='qq_out';\"><img src=\"http://wpa.qq.com/pa?p=1:"+QQNumber+":45\"><a href=\"http://wpa.qq.com/msgrd?v=3&uin="+QQNumber+"&site="+QQName+"&menu=yes\" target=\"_blank\">"+QQName+"</a></p>";
		}
	}
	var WangName,WangNumber;
	if (WangMeter!=""){
		WangMeter=WangMeter.split("|");
		for (var i=0;i<WangMeter.length;i++){
			var WangNumber=WangMeter[i].split(",")[0];
			var WangName=WangMeter[i].split(",")[1];
			Html+="<p class=\"qq_out\" onmouseover=\"this.className='qq_over';\" onmouseout=\"this.className='qq_out';\"><img src=\"Images/Wang.jpg\" /><a href=\"http://amos.im.alisoft.com/msg.aw?v=2&uid="+WangNumber+"&site=cnalichn&s=4\" target=\"_blank\">"+WangName+"</a></p>";
		}
	}
	var MSNName,MSNNumber;
	if (MSNMeter!=""){
		MSNMeter=MSNMeter.split("|");
		for (var i=0;i<MSNMeter.length;i++){
			var MSNNumber=MSNMeter[i].split(",")[0];
			var MSNName=MSNMeter[i].split(",")[1];
			Html+="<p class=\"qq_out\" onmouseover=\"this.className='qq_over';\" onmouseout=\"this.className='qq_out';\"><img src=\"Images/MSN.jpg\" /><a href=\"msnim:chat?contact="+MSNNumber+"\" target=\"_blank\">"+MSNName+"</a></p>";
		}
	}
	var SkypeName,SkypeNumber;
	if (SkypeMeter!=""){
		SkypeMeter=SkypeMeter.split("|");
		for (var i=0;i<SkypeMeter.length;i++){
			var SkypeNumber=SkypeMeter[i].split(",")[0];
			var SkypeName=SkypeMeter[i].split(",")[1];
			Html+="<p class=\"qq_out\" onmouseover=\"this.className='qq_over';\" onmouseout=\"this.className='qq_out';\"><img src=\"Images/Skype.jpg\" /><a href=\"callto://"+SkypeNumber+"\" target=\"_blank\">"+SkypeName+"</a></p>";
		}
	}
	Html+='</div>';
	Html+='<p id=\"qq_online_bottom\"></p>';
	Div.innerHTML=Html;
    //$2("main").appendChild(Div);
	document.body.appendChild(Div);
    FloatDiv("qq_online",160);
}

function FloatDiv(Obj,Ch){
	var Did=$2(Obj);
	var DidTop=parseInt(Did.style.top);
	var Diff=(document.documentElement.scrollTop + Ch - DidTop)*.80;
	Did.style.top=Ch+document.documentElement.scrollTop-Diff+"px";
	FloatID=setTimeout("FloatDiv('"+Obj+"',"+Ch+")",20);
}

function ReCode(ObjID){
	ObjID.src="Inc/GetCode.asp?Meter="+Math.random();
}