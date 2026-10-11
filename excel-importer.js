/* Importador seguro de listas escolares Excel. Solo administrador. */
let pendingExcelImport = null;
function importNorm(v){return String(v??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().replace(/[^A-Z0-9]+/g," ").trim().replace(/\s+/g," ")}
function importCenter(raw){
  const n=importNorm(raw).replace(/^CENTRO ESCOLAR /,"");
  const aliases={"NUEVA ALIANZA":"Nueva Alianza","HUGO CHAVEZ FRIAS":"Hugo Chávez Frías","ROSA SARMIENTO":"Rosa Sarmiento","LA ESPERANZA":"La Esperanza","LA UNION N 3":"La Unión","LA UNION NO 3":"La Unión","LA UNION":"La Unión","SANTA LUCIA":"Santa Lucía","SECTOR DE LIMON":"Sector de Limón","NUEVA JERUSALEN":"Nueva Jerusalén"};
  const name=aliases[n]||n; return centers.find(c=>importNorm(c.nombre)===importNorm(name));
}
function importGrade(sheet,raw){
  const n=importNorm(raw); let name="";
  if(sheet==="PREESCOLAR")name=({"PRIMERO":"Inicial I","SEGUNDO":"Inicial II","TERCERO":"Inicial III"})[n]||"";
  if(sheet==="PRIMARIA")name=({"PRIMERO":"1ro","SEGUNDO":"2do","TERCERO":"3ro","CUARTO":"4to","QUINTO":"5to","SEXTO":"6to"})[n]||"";
  if(sheet==="SECUNDARIA EN EL CAMPO")name=({"SEPTIMO":"7mo","OCTAVO":"8vo","NOVENO":"9no","DECIMO":"10mo","UNDECIMO":"11mo"})[n]||"";
  return grades.find(g=>importNorm(g.nombre)===importNorm(name));
}
function importStatus(sheet,raw){if(sheet==="PREESCOLAR")return "Activo";const n=importNorm(raw);if(n.startsWith("RETIRADO"))return "Retirado";if(n.startsWith("TRASLADADO"))return "Trasladado";if(n.startsWith("EGRESADO"))return "Egresado";return "Activo"}
function readExcelStudents(wb){
  const out=[],warnings=[],counts={};
  for(const sn of ["PREESCOLAR","PRIMARIA","SECUNDARIA EN EL CAMPO"]){
    const ws=wb.Sheets[sn]; if(!ws){warnings.push("Falta la hoja "+sn);continue}
    const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:false});let n=0;
    for(let i=1;i<rows.length;i++){
      const r=rows[i];if(!r||!r.some(v=>String(v??"").trim()))continue;
      const c=importCenter(r[1]),g=importGrade(sn,r[3]);
      const nombres=[r[6],r[7]].map(v=>String(v??"").trim()).filter(Boolean).join(" ");
      const apellidos=[r[8],r[9]].map(v=>String(v??"").trim()).filter(Boolean).join(" ");
      if(!c||!g||!nombres||!apellidos){warnings.push(sn+" fila "+(i+1)+": centro, grado o nombre no reconocido");continue}
      const rawCode=String(r[10]??"").trim().replace(/^"+|"+$/g,"");
      const code=rawCode&&importNorm(rawCode)!=="N A"?rawCode:"SIN-COD-"+String(r[0]||"CENTRO")+"-"+sn.replace(/[^A-Z]/g,"").slice(0,3)+"-"+(i+1);
      const t=importNorm(r[5]);const turno=t==="SABATINO"?"Sabatino":t==="VESPERTINO"?"Vespertino":t==="NOCTURNO"?"Nocturno":"Matutino";
      out.push({centro_id:c.id,codigo_estudiante:code,nombres,apellidos,grado_id:g.id,seccion:String(r[4]??"").trim()||null,turno,estado:importStatus(sn,r[11]),fecha_matricula:new Date().toISOString().slice(0,10),observaciones:"Importado desde Excel: "+sn});n++;
    } counts[sn]=n;
  }
  return {rows:out,warnings,counts};
}
async function previewExcelImport(){
  pendingExcelImport=null;$("runExcelImport").disabled=true;$("excelImportResult").textContent="";
  const file=$("excelImportFile").files?.[0];if(!file){$("excelImportPreview").textContent="Seleccione primero el archivo Excel.";return}
  if(!window.XLSX){$("excelImportPreview").textContent="No cargó el lector Excel. Compruebe Internet y recargue.";return}
  try{const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});pendingExcelImport=readExcelStudents(wb);
    $("excelImportPreview").textContent="Revisión: "+pendingExcelImport.rows.length+" filas válidas. "+Object.entries(pendingExcelImport.counts).map(([k,v])=>k+": "+v).join(" · ")+". PRIMARIA COMUNITARIA se omite sin leer ni modificar. Filas para revisar: "+pendingExcelImport.warnings.length+". Todavía no se ha guardado ningún estudiante.";
    $("runExcelImport").disabled=!pendingExcelImport.rows.length;
  }catch(e){pendingExcelImport=null;$("excelImportPreview").textContent="No se pudo leer el archivo: "+(e?.message||String(e))}
}
async function runExcelImport(){
  if(profile?.rol!=="administrador"){alert("Solo el administrador puede importar listas.");return}
  if(!navigator.onLine){alert("Conéctese a Internet antes de importar.");return}
  if(!pendingExcelImport?.rows?.length){alert("Primero seleccione y revise el Excel.");return}
  const rows=pendingExcelImport.rows;
  if(!confirm("Se revisaron "+rows.length+" filas. Se omitirán duplicados y no se reemplazará ningún registro existente. PRIMARIA COMUNITARIA no se importará. ¿Continuar?"))return;
  $("runExcelImport").disabled=true;$("excelImportResult").textContent="Comprobando estudiantes existentes…";
  let inserted=0,duplicates=0,failed=0;const errors=[];
  try{
    const q=await client.from("estudiantes").select("centro_id,codigo_estudiante,nombres,apellidos");if(q.error)throw q.error;
    const codes=new Set((q.data||[]).map(x=>x.centro_id+"|"+importNorm(x.codigo_estudiante)));
    const names=new Set((q.data||[]).map(x=>x.centro_id+"|"+importNorm(x.nombres)+"|"+importNorm(x.apellidos)));
    const unique=[];
    for(const r of rows){const ck=r.centro_id+"|"+importNorm(r.codigo_estudiante),nk=r.centro_id+"|"+importNorm(r.nombres)+"|"+importNorm(r.apellidos);
      if(codes.has(ck)||names.has(nk)){duplicates++;continue}codes.add(ck);names.add(nk);unique.push(r)}
    for(let i=0;i<unique.length;i+=40){const b=unique.slice(i,i+40),res=await client.from("estudiantes").insert(b);
      if(res.error){failed+=b.length;errors.push("Lote "+(Math.floor(i/40)+1)+": "+res.error.message)}else inserted+=b.length;
      $("excelImportResult").textContent="Avance: "+inserted+" agregados; "+duplicates+" duplicados omitidos; "+failed+" con error."}
    await loadStudents();await loadStats();
    $("excelImportResult").textContent="IMPORTACIÓN FINALIZADA\nAgregados: "+inserted+"\nDuplicados omitidos sin alterar registros: "+duplicates+"\nFilas con error: "+failed+"\nFilas para revisar: "+pendingExcelImport.warnings.length+(errors.length?"\n"+errors.join("\n"):"")+(pendingExcelImport.warnings.length?"\nAvisos iniciales:\n"+pendingExcelImport.warnings.slice(0,10).join("\n"):"");
    $("excelImportPreview").textContent="Proceso terminado. Revise el resultado antes de volver a cargar el archivo.";pendingExcelImport=null;
  }catch(e){$("excelImportResult").textContent="La importación se detuvo: "+(e?.message||String(e))+"\nLos registros existentes no se reemplazan."}
  finally{$("runExcelImport").disabled=!pendingExcelImport?.rows?.length}
}
$("previewExcelImport").onclick=previewExcelImport;
$("excelImportFile").onchange=()=>{pendingExcelImport=null;$("runExcelImport").disabled=true;$("excelImportPreview").textContent="Archivo seleccionado. Pulse «Revisar archivo».";$("excelImportResult").textContent=""};
$("runExcelImport").onclick=runExcelImport;
