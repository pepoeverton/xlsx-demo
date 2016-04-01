;(function(){
	'use strict';

	var upload = document.querySelector('.input-file');
	var body = document.querySelector('#table-body');
	var head = document.querySelector('#table-head');
	var X = XLSX;
	var ssf = SSF;
	var output = {};

	function app(){
		upload.addEventListener('change', handleFile, false);
	}

	function toJson(workbook) {
	  var result = {};
	  workbook.SheetNames.forEach(function(sheetName) {
	    var roa = X.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
	    if(roa.length > 0){
	      result[sheetName] = roa;
	    }
	  });
	  return result;
	}

	function fixData(data) {
	  var o = "", l = 0, w = 10240;
	  for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
	  o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
	  return o;
	}

	function process(body) {
	  output = JSON.stringify(toJson(body), 2, 2);
	  var json = JSON.parse(output);
	  if(typeof console !== 'undefined') console.log("output: ", new Date());
		render(json);
	}

	function handleFile(e) {
	  var files = e.target.files;
	  var f = files[0];
	  {
	    var reader = new FileReader();
	    var name = f.name;
	    reader.onload = function(e) {
	      if(typeof console !== 'undefined') console.log("input: ", new Date());
	      var data = e.target.result;
	      var arr = fixData(data);
	      var body;
	      body = X.read(btoa(arr), {type: 'base64'});
	      process(body);
	    };
	    reader.readAsArrayBuffer(f);
	  }
	}
	function render(json){
		var plan = json[Object.keys(json)[0]];
		renderHead(plan[0]);
		var tr = document.createDocumentFragment();
		for(var value in plan){
			var valu = plan[value];
			var linha = document.createElement('tr');
			for(var prop in valu){
				var td = document.createElement('td');
				var span = document.createElement('span');
				span.appendChild(document.createTextNode(valu[prop]));
				td.appendChild(span);
				linha.appendChild(td);
			}
			tr.appendChild(linha);
		}
		body.appendChild(tr);
	}
	function renderHead(json){
		var linha = document.createElement('tr');
		for(var i in json){
			var th = document.createElement('th');
			var span = document.createElement('span');
			span.appendChild(document.createTextNode(i));
			th.appendChild(span);
			linha.appendChild(th);
		}
		head.appendChild(linha);
	}

	window.addEventListener('DOMContentLoaded', app);
})();
