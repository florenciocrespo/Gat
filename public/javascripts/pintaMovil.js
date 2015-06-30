var table = document.createElement('table');
var thead = document.createElement('thead');
var tbody  = document.createElement('tbody'); 
table.setAttribute("id", "facMovil");
var td = document.createElement('th');
var tr = document.createElement('tr');

function lineaNueva(){
	tbody.appendChild(tr);
  	table.appendChild(tbody); 
}

function blanco(){
	td = document.createElement('td');
    td.innerHTML=""; 
    tr.appendChild(td);
}
function blancoTh(){
	td = document.createElement('th');
    td.innerHTML=""; 
    tr.appendChild(td);
}