'Imports System.IO
'Imports System.Text
'Imports System.Web.Script.Serialization
'Imports VBA_CORE.modTypes
'Imports VBA_CORE.Logger

'Public Module modReports

'    Public Sub GenerateReports(exported As List(Of ExportedModule))
'        Try
'            Dim serializer As New JavaScriptSerializer()
'            serializer.MaxJsonLength = Integer.MaxValue

'            Dim outJson = Path.Combine(EXPORT_DIR, "FunctionsIndex.json")
'            File.WriteAllText(outJson, serializer.Serialize(exported), Encoding.UTF8)

'            Dim txt = Path.Combine(EXPORT_DIR, "FunctionsIndex.txt")
'            Using sw As New StreamWriter(txt, False, Encoding.UTF8)
'                For Each m In exported
'                    sw.WriteLine($"[{m.Type}] {m.Name}")
'                    If m.Functions.Count > 0 Then
'                        For Each f In m.Functions
'                            sw.WriteLine("   " & f.Signature)
'                            sw.WriteLine($"      Lines: {f.LineCount}, Complexity: {f.Complexity}, Access: {f.Accessibility}, CalledBy: {f.CalledBy.Count}, ParentClass: {If(String.IsNullOrEmpty(f.ParentClass), "-", f.ParentClass)}")
'                        Next
'                    End If
'                    If m.References.Count > 0 Then
'                        sw.WriteLine("   Referit de:")
'                        For Each r In m.References.Distinct()
'                            sw.WriteLine("      " & r)
'                        Next
'                    End If
'                    sw.WriteLine()
'                Next
'            End Using
'            LogInfo("✅ FunctionsIndex.json / .txt generate.")
'        Catch ex As Exception
'            LogError("GenerateReports", ex)
'        End Try
'    End Sub

'    Public Sub GenerateHTML(exported As List(Of ExportedModule))
'        Try
'            Dim htmlPath = Path.Combine(EXPORT_DIR, "FunctionsIndex.html")
'            Using html As New StreamWriter(htmlPath, False, Encoding.UTF8)
'                Dim serializer As New JavaScriptSerializer()
'                serializer.MaxJsonLength = Integer.MaxValue

'                Dim dataObj = exported.Select(Function(m) New With {
'                    .Name = m.Name, .Type = m.Type,
'                    .Functions = m.Functions.Select(Function(f) New With {
'                        .Name = f.Name,
'                        .Signature = f.Signature,
'                        .Lines = f.LineCount,
'                        .Complexity = f.Complexity,
'                        .Access = f.Accessibility,
'                        .CalledBy = f.CalledBy.Count,
'                        .ParentClass = f.ParentClass
'                    }).ToList(),
'                    .Incoming = m.IncomingModules.ToList(),
'                    .Outgoing = m.OutgoingModules.ToList(),
'                    .Refs = m.References.Distinct().ToList()
'                }).ToList()

'                Dim dataJson = serializer.Serialize(dataObj)

'                html.WriteLine("<!DOCTYPE html><html><head><meta charset='utf-8'/>")
'                html.WriteLine("<title>Access VBA Project Map</title>")
'                html.WriteLine("<meta http-equiv='X-UA-Compatible' content='IE=edge'/>")
'                html.WriteLine("<style>")
'                html.WriteLine("body{font-family:Segoe UI,Arial;margin:20px;background:#f6f8fa}")
'                html.WriteLine("#wrap{display:grid;grid-template-columns:1fr 360px;grid-gap:16px}")
'                html.WriteLine("#network{width:100%;height:720px;border:1px solid #aaa;border-radius:8px;background:#fff}")
'                html.WriteLine("#side{background:#fff;border-radius:8px;box-shadow:0 1px 4px #ccc;padding:12px;height:720px;overflow:auto}")
'                html.WriteLine(".legend span{display:inline-block;margin-right:10px;padding:4px 8px;border-radius:6px}")
'                html.WriteLine("input[type=checkbox]{margin-right:6px}")
'                html.WriteLine(".mod{margin:10px 0;padding:8px;border-bottom:1px dashed #ddd}")
'                html.WriteLine(".func{margin-left:15px}")
'                html.WriteLine(".muted{color:#777}")
'                html.WriteLine("</style>")
'                html.WriteLine("<script src='https://cdn.jsdelivr.net/npm/vis-network@9.1.2/dist/vis-network.min.js'></script>")
'                html.WriteLine("</head><body>")
'                html.WriteLine("<h1>Access VBA Project Map</h1>")
'                html.WriteLine("<div class='legend'>")
'                html.WriteLine("<label><input type='checkbox' id='flModules' checked>Modules</label>")
'                html.WriteLine("<label><input type='checkbox' id='flForms' checked>Forms</label>")
'                html.WriteLine("<label><input type='checkbox' id='flReports' checked>Reports</label>")
'                html.WriteLine("<label><input type='checkbox' id='flMacros' checked>Macros</label>")
'                html.WriteLine("</div>")
'                html.WriteLine("<div id='wrap'><div id='network'></div><div id='side'><b>Detalii</b><div id='details' style='margin-top:10px;color:#333'>Selectează un nod...</div><hr/>")
'                html.WriteLine("<input id='search' placeholder='Caută modul/funcție...' style='width:100%;padding:6px;margin:10px 0'/>")
'                html.WriteLine("<div id='list'></div></div></div>")

'                html.WriteLine("<script>")
'                html.WriteLine("document.addEventListener('DOMContentLoaded', function(){")
'                html.WriteLine("var data=" & dataJson & ";")
'                html.WriteLine("function colorOf(t){return {Modules:'#007bff',Forms:'#28a745',Reports:'#ffc107',Macros:'#6f42c1'}[t]||'#999';}")
'                html.WriteLine("var nodesArr=[],edgesArr=[];")

'                html.WriteLine("data.forEach(function(m){")
'                html.WriteLine("  var nodeId = m.Type + ':' + m.Name;")
'                html.WriteLine("  nodesArr.push({id:nodeId,label:m.Name,group:m.Type,color:colorOf(m.Type)});")
'                html.WriteLine("  (m.Outgoing||[]).forEach(function(o){")
'                html.WriteLine("     var target = data.find(x=>x.Name===o)||{Type:'',Name:o};")
'                html.WriteLine("     var targetId = target.Type ? (target.Type+':'+target.Name) : o;")
'                html.WriteLine("     edgesArr.push({from:nodeId,to:targetId,arrows:'to'});")
'                html.WriteLine("  });")
'                html.WriteLine("});")

'                html.WriteLine("var nodesDS=new vis.DataSet(nodesArr);")
'                html.WriteLine("var edgesDS=new vis.DataSet(edgesArr);")

'                html.WriteLine("var network=new vis.Network(document.getElementById('network'),{nodes:nodesDS,edges:edgesDS},{")
'                html.WriteLine("edges:{arrows:{to:{enabled:true,scaleFactor:0.8}}},")
'                html.WriteLine("physics:{enabled:true,barnesHut:{gravitationalConstant:-8000,springLength:250},stabilization:{iterations:250}},")
'                html.WriteLine("interaction:{hover:true}")
'                html.WriteLine("});")

'                html.WriteLine("network.once('stabilizationIterationsDone',function(){network.fit({animation:{duration:800}});});")

'                html.WriteLine("function applyFilters(){var show={Modules:flModules.checked,Forms:flForms.checked,Reports:flReports.checked,Macros:flMacros.checked};")
'                html.WriteLine(" nodesDS.get().forEach(function(n){var hidden=!show[n.group];nodesDS.update({id:n.id,hidden:hidden});});}")
'                html.WriteLine("['flModules','flForms','flReports','flMacros'].forEach(id=>document.getElementById(id).addEventListener('change',applyFilters));")

'                html.WriteLine("network.on('click',function(params){")
'                html.WriteLine(" if(params.nodes&&params.nodes.length){var id=params.nodes[0];var m=data.find(x=>(x.Type+':'+x.Name)===id);")
'                html.WriteLine("   if(!m){details.innerHTML='';return;}")
'                html.WriteLine("   var html='<h3>['+m.Type+'] '+m.Name+'</h3>';")
'                html.WriteLine("   html+='<div class=""muted""><b>Funcții:</b> '+m.Functions.length+'</div><ul>';")
'                html.WriteLine("   m.Functions.forEach(function(f){html+='<li>'+f.Signature+' <span class=""muted"">(lines:'+f.Lines+', cplx:'+f.Complexity+', access:'+f.Access+', calledBy:'+f.CalledBy+', class:'+(f.ParentClass||'-')+')</span></li>';});")
'                html.WriteLine("   html+='</ul><b>Primește apeluri de la:</b><ul>'+(m.Incoming||[]).map(i=>'<li>'+i+'</li>').join('')+'</ul>';")
'                html.WriteLine("   details.innerHTML=html;")
'                html.WriteLine(" }else{details.innerHTML='Selectează un nod...';}});")

'                html.WriteLine("function renderList(q){q=(q||'').toLowerCase();var out='';")
'                html.WriteLine(" data.forEach(function(m){var showMod=m.Name.toLowerCase().includes(q);")
'                html.WriteLine(" var ff=(m.Functions||[]).filter(f=>showMod||(f.Signature||'').toLowerCase().includes(q)||(f.Name||'').toLowerCase().includes(q));")
'                html.WriteLine(" if(showMod||ff.length){out+='<div class=""Mod""><b>['+m.Type+'] '+m.Name+'</b>';")
'                html.WriteLine(" if(ff.length){out+='<ul>'+ff.map(f=>'<li class=""func"">'+f.Signature+' <span class=""muted"">(lines:'+f.Lines+', cplx:'+f.Complexity+', access:'+f.Access+')</span></li>').join('')+'</ul>';}")
'                html.WriteLine(" out+='</div>';}});")
'                html.WriteLine(" list.innerHTML=out;}")
'                html.WriteLine("search.addEventListener('keyup',()=>renderList(search.value));")
'                html.WriteLine("renderList('');applyFilters();")
'                html.WriteLine("console.log('Project Map loaded:',{nodes:nodesDS.length,edges:edgesDS.length});")
'                html.WriteLine("});")
'                html.WriteLine("</script></body></html>")
'            End Using
'            LogInfo("✅ FunctionsIndex.html generat.")
'        Catch ex As Exception
'            LogError("GenerateHTML", ex)
'        End Try
'    End Sub

'    Public Sub ExportAIReady(exported As List(Of ExportedModule))
'        Try
'            Dim compact = exported.Select(Function(m) New With {
'                .Name = m.Name,
'                .Type = m.Type,
'                .Functions = m.Functions.Select(Function(f) New With {.Name = f.Name, .Sig = f.Signature, .Lines = f.LineCount, .Cplx = f.Complexity, .Access = f.Accessibility, .ParentClass = f.ParentClass}).ToList(),
'                .Incoming = m.IncomingModules.ToList(),
'                .Outgoing = m.OutgoingModules.ToList()
'            }).ToList()
'            Dim serializer As New JavaScriptSerializer()
'            serializer.MaxJsonLength = Integer.MaxValue
'            File.WriteAllText(Path.Combine(EXPORT_DIR, "Functions_AI.json"), serializer.Serialize(compact), Encoding.UTF8)

'            Dim sb As New StringBuilder()
'            For Each p In fileCache.Keys.OrderBy(Function(x) x)
'                sb.AppendLine("''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''")
'                sb.AppendLine("FILE: " & p.Replace(EXPORT_DIR, "").TrimStart("\"c))
'                sb.AppendLine("''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''")
'                Try
'                    sb.AppendLine(fileCache(p).Content)
'                Catch ex As Exception
'                    LogError($"ExportAIReady Read({Path.GetFileName(p)})", ex)
'                End Try
'                sb.AppendLine()
'            Next
'            File.WriteAllText(Path.Combine(EXPORT_DIR, "Code_AI.txt"), sb.ToString(), Encoding.UTF8)
'            LogInfo("✅ Export AI-ready: Functions_AI.json + Code_AI.txt")
'        Catch ex As Exception
'            LogError("ExportAIReady", ex)
'        End Try
'    End Sub

'End Module
