 $(function() {

 	$("div").dblclick(function() {
 		var h=$(this).height()  
 		var setH=20
 		if(h!=setH){
 			$(this).attr("style","height:"+setH+"px;overflow:hidden;background-color:#999999;border-bottom:1px solid #000;")
 		}else{
 			$(this).attr("style","")
 		}
 	})
 })