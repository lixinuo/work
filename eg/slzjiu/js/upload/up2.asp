<script type="text/javascript" src="../lanyu/js/anplus.js"></script>
<script type="text/javascript" src="../lanyu/js/AjaxUploader2.js"></script>
<script type="text/javascript">
var AjaxUp=null;
window.onload=function(){
	_.EndragEx("uploader_Title","upload_Box",0,0);
	
};

function showUpload(obj, inputID){
	try{
		AjaxUp.reset();
	}catch(ex){}
	var ps=_.abs(obj);
	var box=_.$("upload_Box");
	box.style.left=ps.x - parseInt(box.style.width)/2 + 30 + "px";
	box.style.top=ps.y +25 + "px";
	box.style.display="block";
	//创建Uploader，参数Contenter	字符串	上传控件的容器，程序自动给容易四面增加3px的padding
	AjaxUp=new AjaxProcesser("uploadContenter");
	
	//设置提交到的iframe名称
	AjaxUp.target="AnUploader";  
	
	//上传处理页面
	AjaxUp.url="../lanyu/upload.asp"; 
	
	//保存目录
	AjaxUp.savePath="_upload";  
	var e=1;
	//上传成功时运行的程序
	AjaxUp.succeed=function(files){
		//下面遍历所有的文件，files是一个数组，数组元素的数目就是上传文件的个数，每个元素包含的信息为文件名字和文件大小
		var info="";
		for(var i=0;i<files.length;i++){
			info+=files[i].name + ";";
		}
		//info=_.$(inputID).value+info.substr(0,info.length-1)+";";	
		clearTimeout(saved);
		var _ups;
		if (e==1) {
		_ups=setTimeout(function(){							   
							   _.$(inputID).value=_.$(inputID).value+info.substr(0,info.length-1)+";";					   
							   							   
							   },500);	
		}else{
			return false;
			//$("#files").append(info.substr(0,info.length-1)+";");
		}
		
		
		//_.$(inputID).value+=info; 
		box.style.display="none";
		e++;
		return false;
		
	}
	clearTimeout(_ups);
	//上传失败时运行的程序
	AjaxUp.faild=function(msg){
		alert("失败原因:" + msg)
	}
	
}
</script>  
