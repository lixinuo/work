
 function checkAll(form,obj)
 {
    for (var i = 0; i < form.elements.length; i++) {
        var e = form.elements[i];
        if(e.disabled!='disabled'&&e.name=='selID'&&e.type=='checkbox')
        e.checked = obj.checked;
    }
}
 function checkFan()
 {
	 var form= document.forms[0];
	 for (var i = 0; i < form.elements.length; i++) {
        var e = form.elements[i];
        if(e.disabled!='disabled'&&e.name=='selID'&&e.type=='checkbox')
        { e.checked = !e.checked;
		}
    }
 }
