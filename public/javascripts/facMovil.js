var precio=0.0069;
var emision=0.0042;
var limiteTmoDatos=393;
var limiteTmoServicios=300;
var limiteTmoTerminales=350;
var limiteTmoCT=1800;
var limiteTmoSE=3872;
var costeHSEDiurno=0.00663;var costeHSEResto=0.00730;
var costeHComunidadDiurno=0.00488;var costeHComunidadResto=0.00537;
var costeHPersonalizadaDiurno=0.00488;var costeHPersonalizadaResto=0.00537;
var costeHNocturno=0.00601;
//var chat=['03','10','17','24','31'];//abril '02','03','05','12','19','26'
var mesPremium=['01','03','10','17','24','31']//abril '02','03','05','12','19','23','26'
var totalFac=0;
var totalFac1=0;
var totalFac2=0;
var insatisfechos=1700;
var insatisfechosSMD=2100;
var buzonClasico=1631;
var buzonClasicoReset=3421;
var buzonSMD=358;
var tmoInsatisfechos=201;
var tmoInsatisfechosSMD=263;

$(function(){
  $("#excel").on('click', ExportHTMLTableToExcel);

});
 ///

function fileInfo(e){
    var file = e.target.files[0];
    if (file.name.split(".")[1].toUpperCase() != "CSV"){
      alert('Invalid csv file !');
      e.target.parentNode.reset();
      return;
    }else{
      document.getElementById('file_info').innerHTML = "<p>Mes: "+file.name + " | "+file.size+" Bytes.</p>";
    }
  }
 function handleFileSelect(){
    var contadorDatos=0,contadorServicios=0,contadorTerminales=0,
   contadorChat=0,contadorPersonalizada=0,contadorComunidad=0,contadorNocturno=0,
   contadorProyectos=0; 
   var tmoDatos=0,tmoServicios=0,tmoTerminales=0,
   tmoChat=0,tmoPersonalizada=0,tmoComunidad=0,tmoNocturno=0,
   tmoProyectos=0; 
   tr = document.createElement('tr');
    td = document.createElement('th');
    td.innerHTML ="unidades";
    tr.appendChild(td);
    td = document.createElement('th');
    td.innerHTML ="atendidas";
    tr.appendChild(td); 
    td = document.createElement('th');
    td.innerHTML ="TMO";
    tr.appendChild(td); 
    td = document.createElement('th');
    td.innerHTML ="€";
    tr.appendChild(td);  
    td = document.createElement('th');
    td.innerHTML ="segundos diurnos";
    tr.appendChild(td);  
    td = document.createElement('th');
    td.innerHTML ="segundos resto";
    tr.appendChild(td);        
    lineaNueva(); 
    var file = document.getElementById("the_file").files[0];
    var reader = new FileReader();
    var link_reg = /(http:\/\/|https:\/\/)/i;
    reader.onload = function(file) {
    var content = file.target.result;
    var rows = file.target.result.split(/[\r\n|\n]+/);
  
  for (var i = 0; i < rows.length; i++){
    var arr = rows[i].split(';');//delimitador
    for (var j = 0; j < arr.length; j++){
       switch (j) { 
                  case 0: dia = arr[j].replace(/\"/g,'');                  
                  break
                  case 4: unidad = arr[j].replace(/\"/g,'');                  
                  break
                  case 8: hora = arr[j].replace(/\"/g,'');                  
                  break
                  case 11: atendidas = +arr[j].replace(/\./g,'');                 
                  break 
                  case 12: atendidas20 = +arr[j].replace(/\./g,'');                 
                  break
                  case 13: tmo1 = +arr[j].replace(/\./g,'');                 
                  break  
                  case 14: tmo2 = +arr[j].replace(/\./g,'');                 
                  break  
                  case 15: tmo = +arr[j].replace(/\./g,'');                 
                  break   
                  case 16: Staff = +arr[j].replace(/\./g,'');  
                           if(unidad.indexOf('CAT Datos') !=-1) {
                              contadorDatos+=atendidas;
                              tmoDatos+=tmo+tmo1+tmo2;
                           }else if (unidad.indexOf('CAT Servicios') !=-1) {
                              contadorServicios+=atendidas;
                              tmoServicios+=tmo+tmo1+tmo2;
                           }else if(unidad.indexOf('CAT Terminales') !=-1) {
                              contadorTerminales+=atendidas;
                              tmoTerminales+=tmo+tmo1+tmo2;
                           }else if(unidad.indexOf('Chat') !=-1) {
                             //var h= +hora.substring(3,5) >= 22 ? true : 
                             //       +hora.substring(3,5) < 08 ? true : false;
                             var d= dia.substring(0,2);                               
                             if(mesPremium.indexOf(d) == -1 ){
                                     tmoChat+=Staff; 
                               }else{
                                     contadorChat+=Staff;                                     
                                }
                           }else if(unidad.indexOf('Comunidad') !=-1) {
                              var h= +hora.substring(3,5) >= 22 ? true : 
                                    +hora.substring(3,5) < 08 ? true : false;
                              var d= dia.substring(0,2); 
                             if(mesPremium.indexOf(d) == -1 && !h){
                                 tmoComunidad+=Staff;                               
                             }else{contadorComunidad+=Staff; }
                             
                           }else if (unidad.indexOf('Personalizada') !=-1) {
                               var h= +hora.substring(3,5) >= 22 ? true : 
                                    +hora.substring(3,5) < 08 ? true : false;
                              var d= dia.substring(0,2); 
                             if(mesPremium.indexOf(d) == -1 && !h){
                                 tmoPersonalizada+=Staff;                               
                             }else{contadorPersonalizada+=Staff;}
                           }else if(unidad.indexOf('MoviStar Contrato Nocturno') !=-1) {
                              contadorNocturno+=Staff;
                             
                           }else if (unidad.indexOf('Proyectos') !=-1) {

                             var h= +hora.substring(3,5) >= 22 ? true :  
                                    +hora.substring(3,5) < 08 ? true : false;
                             var d= dia.substring(0,2); 
                             if(mesPremium.indexOf(d) == -1 && !h){
                                 tmoProyectos+=Staff; 

                                                              
                             }else{
                                  contadorProyectos+=Staff;
                                  //console.log("hora: "+hora.substring(3,5));
                                 
                                 
                             }
                             
                       

                            
                           }                                                                              
                  break 

       }//switch
       
    }//rof

  }//rof

tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="CAT Servicios Básicos de Voz";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(contadorServicios, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
var tmo=(tmoServicios/contadorServicios) > limiteTmoServicios ? limiteTmoServicios : (tmoServicios/contadorServicios);
td.innerHTML=formato_numero(tmo, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.round(tmo)*precio*contadorServicios, 2, ",", ".");
totalFac+=Math.round(tmo)*precio*contadorServicios;
tr.appendChild(td);
blanco();blanco();
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Servicios Móviles Datos";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(contadorDatos, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
var tmo=(tmoDatos/contadorDatos) > limiteTmoDatos ? limiteTmoDatos : (tmoDatos/contadorDatos);
td.innerHTML=formato_numero(tmo, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.round(tmo)*precio*contadorDatos, 2, ",", ".");
totalFac+=Math.round(tmo)*precio*contadorDatos;
tr.appendChild(td);
blanco();blanco();
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="CAT Terminales";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(contadorTerminales, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
var tmo=(tmoTerminales/contadorTerminales) > limiteTmoTerminales ? limiteTmoTerminales : (tmoTerminales/contadorTerminales);
td.innerHTML=formato_numero(tmo, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(tmo*precio*contadorTerminales, 2, ",", ".");
totalFac+=tmo*precio*contadorTerminales;
tr.appendChild(td);
blanco();blanco();
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Total";
tr.appendChild(td);
blanco();
blanco();
td = document.createElement('td');
td.innerHTML=formato_numero(totalFac, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
lineaNueva();
/////////////////////////////////////////
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Premium";
tr.appendChild(td);
td = document.createElement('td');
var aux1=Math.round(tmoProyectos*costeHSEDiurno);
var aux2=Math.round(contadorProyectos*costeHSEResto);
var aux3=aux1+aux2;
aux3=Math.round(aux3-(aux3*0.05));
td.innerHTML= formato_numero(aux3/(precio*limiteTmoSE), 0, ",", ".");//(aux1*costeHSEDiurno)+(aux2*costeHSEResto)
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(limiteTmoSE, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(aux3, 2, ",", ".");
totalFac1+=aux3;
tr.appendChild(td);
//var hours = Math.floor( tmoProyectos / 3600 );  
//var minutes = Math.floor( (tmoProyectos % 3600) / 60 );
//console.log("aux1 "+aux1+" aux2 "+aux2+" aux3 "+aux3);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.round(tmoProyectos), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.round(contadorProyectos), 0, ",", ".");
tr.appendChild(td);
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Soporte Att. Nocturna";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero((contadorNocturno*costeHNocturno)/(precio*limiteTmoCT), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(limiteTmoCT, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(contadorNocturno*costeHNocturno, 2, ",", ".");
tr.appendChild(td);
totalFac1+=contadorNocturno*costeHNocturno;
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(contadorNocturno ), 0, ",", ".");
tr.appendChild(td);
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="CT Comunidad";
tr.appendChild(td);
td = document.createElement('td');
var aux1=tmoComunidad*costeHComunidadDiurno;
var aux2=contadorComunidad*costeHComunidadResto;
var aux3=aux1+aux2;
td.innerHTML= formato_numero(aux3/(precio*limiteTmoCT), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(limiteTmoCT, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero((aux3/(precio*limiteTmoCT))*precio*limiteTmoCT, 2, ",", ".");
totalFac1+=(aux3/(precio*limiteTmoCT))*precio*limiteTmoCT;
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(tmoComunidad), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(contadorComunidad), 0, ",", ".");
tr.appendChild(td);
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="CT ATT. Personalizada";
tr.appendChild(td);
td = document.createElement('td');
var aux1=tmoPersonalizada*costeHPersonalizadaDiurno;
var aux2=contadorPersonalizada*costeHPersonalizadaResto;
var aux3=aux1+aux2;
td.innerHTML= formato_numero(aux3/(precio*limiteTmoCT), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(limiteTmoCT, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero((aux3/(precio*limiteTmoCT))*precio*limiteTmoCT, 2, ",", ".");
totalFac1+=(aux3/(precio*limiteTmoCT))*precio*limiteTmoCT;
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(tmoPersonalizada), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(contadorPersonalizada), 0, ",", ".");
tr.appendChild(td);
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="CT Chat";
tr.appendChild(td);
td = document.createElement('td');
var aux1=tmoChat*costeHComunidadDiurno;
var aux2=contadorChat*costeHComunidadResto;
var aux3=aux1+aux2;
td.innerHTML= formato_numero(aux3/(precio*limiteTmoCT), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(limiteTmoCT, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero((aux3/(precio*limiteTmoCT))*precio*limiteTmoCT, 2, ",", ".");
totalFac1+=(aux3/(precio*limiteTmoCT))*precio*limiteTmoCT;
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(tmoChat), 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(Math.floor(contadorChat), 0, ",", ".");
tr.appendChild(td);
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Total";
tr.appendChild(td);
blanco();
blanco();
td = document.createElement('td');
td.innerHTML=formato_numero(totalFac1, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Insatisfechos";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(insatisfechos, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="201";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(insatisfechos*tmoInsatisfechos*emision, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
totalFac2+=insatisfechos*tmoInsatisfechos*emision;
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Insatisfechos SMD";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(insatisfechosSMD, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="263";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(insatisfechosSMD*tmoInsatisfechosSMD*emision, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
totalFac2+=insatisfechosSMD*tmoInsatisfechosSMD*emision;
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Buzón App Clásicos";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonClasico, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="201";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonClasico*tmoInsatisfechos*emision, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
totalFac2+=buzonClasico*tmoInsatisfechos*emision;
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Buzón App Clásicos Reset";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonClasicoReset, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="201";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonClasicoReset*tmoInsatisfechos*precio, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
totalFac2+=buzonClasicoReset*tmoInsatisfechos*precio;
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Buzón App SMD y Correo";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonSMD, 0, ",", ".");
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML="263";
tr.appendChild(td);
td = document.createElement('td');
td.innerHTML=formato_numero(buzonSMD*tmoInsatisfechosSMD*emision, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
totalFac2+=buzonSMD*tmoInsatisfechosSMD*emision;
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="Total";
tr.appendChild(td);
blanco();
blanco();
td = document.createElement('td');
td.innerHTML=formato_numero(totalFac2, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
lineaNueva();
tr = document.createElement('tr');
td = document.createElement('td');
td.innerHTML="TOTAL";
tr.appendChild(td);
blanco();
blanco();
td = document.createElement('td');
td.innerHTML=formato_numero(totalFac+totalFac1+totalFac2, 2, ",", ".");
tr.appendChild(td);
blanco();blanco();
lineaNueva();


document.getElementById('list').appendChild(table);
          };
  reader.readAsText(file);
 }
 
function formato_numero(numero, decimales, separador_decimal, separador_miles){ // v2007-08-06
    numero=parseFloat(numero);
    if(isNaN(numero)){
        return "";
    }

    if(decimales!==undefined){
        // Redondeamos
        numero=numero.toFixed(decimales);
    }

    // Convertimos el punto en separador_decimal
    numero=numero.toString().replace(".", separador_decimal!==undefined ? separador_decimal : ",");

    if(separador_miles){
        // Añadimos los separadores de miles
        var miles=new RegExp("(-?[0-9]+)([0-9]{3})");
        while(miles.test(numero)) {
            numero=numero.replace(miles, "$1" + separador_miles + "$2");
        }
    }

    return numero;
}

  function ExportHTMLTableToExcel()
    {
       var thisTable = document.getElementById("list").innerHTML;
       window.clipboardData.setData("Text", thisTable);
       var objExcel = new ActiveXObject ("Excel.Application");
       objExcel.visible = true;
       var objWorkbook = objExcel.Workbooks.Add;
       var objWorksheet = objWorkbook.Worksheets(1);
       objWorksheet.Paste;
    }    